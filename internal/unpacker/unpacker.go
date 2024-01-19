package unpacker

import (
	"container/list"
	"encoding/json"
	"fmt"
	"io"
	"log/slog"
	"os"
	"path"
	"path/filepath"

	"github.com/ashdwilson/ole/pkg/models"
	"github.com/ashdwilson/ole/pkg/unpackers"
	"github.com/gabriel-vasile/mimetype"
)

func Unpack(infilePath, outdirPath string) (err error) {
	logger := slog.New(slog.NewJSONHandler(os.Stdout, nil))
	slog.SetDefault(logger)
	parsed := &models.Results{ParsedFiles: map[string]*models.Result{}}
	toBeParsed := list.New()

	// Check that putdirPath exists and is a directory
	outDirStat, err := os.Stat(outdirPath)
	if err != nil {
		err = fmt.Errorf("%w: getting stat on output dir", err)
		return
	}
	if !outDirStat.IsDir() {
		err = fmt.Errorf("output path is not a dir")
		return
	}

	err = copyFileToDir(infilePath, outdirPath)
	if err != nil {
		return
	}
	newInFilePath := path.Join(outdirPath, path.Base(infilePath))
	// Add the input file to the queue
	toBeParsed.PushBack(newInFilePath)

	for toBeParsed.Len() > 0 {
		np := toBeParsed.Front()
		nextPath, ok := np.Value.(string)
		toBeParsed.Remove(np)
		if !ok {
			slog.Error("unable to get value from queue item")
			continue
		}
		err = unpackFile(nextPath, parsed, toBeParsed)
		if err != nil {
			parsed.ParsedFiles[nextPath].Error = err.Error()
		}
	}
	// Write results to file
	var logFile *os.File
	logFile, err = os.Create(path.Join(outdirPath, "ole.log"))
	if err != nil {
		return
	}
	err = json.NewEncoder(logFile).Encode(parsed)
	return

}

// unpackFile unpacks all members from the file (if supported), updates the results struct,
// and adds all new files to the queue.
func unpackFile(fname string, results *models.Results, queue *list.List) (err error) {
	var inFile *os.File
	inFile, err = os.Open(fname)
	if err != nil {
		err = fmt.Errorf("%w: opening input file", err)
		return
	}
	defer inFile.Close()
	var mType *mimetype.MIME
	mType, err = getTypeFromReader(fname, inFile)
	if err != nil {
		err = fmt.Errorf("%w: rewinding readseeker", err)
	}
	fileInfo, err := inFile.Stat()
	if err != nil {
		err = fmt.Errorf("%w: unable to stat file", err)
		return
	}
	fileSize := fileInfo.Size()
	var unpackerImpl unpackers.Unpacker
	mTypeStr := mType.String()
	results.ParsedFiles[fname] = &models.Result{
		FileType:  mTypeStr,
		Expanded:  false,
		Supported: false,
		Error:     "",
	}
	switch mTypeStr {
	// Catch the modern MS office docs
	case "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
		unpackerImpl = &unpackers.OfficeZip{}
		results.ParsedFiles[fname].Supported = true
	// Grab OLEv2, or MS-CFB
	case "application/x-ole-storage":
		unpackerImpl = &unpackers.MSCFB{}
		// This can be avariety of things... including OLE 1.0
	case "application/octet-stream":
		fileName := path.Base(fname)
		fileExtension := filepath.Ext(fileName)
		switch fileName {
		case "Ole10Native":
			results.ParsedFiles[fname].Supported = true
			unpackerImpl = &unpackers.OLE10Native{}
		case "CompObj", "ObjInfo":
			results.ParsedFiles[fname].Supported = true
			return
		}
		switch fileExtension {
		case "emf":
			results.ParsedFiles[fname].Supported = true
			return

		default:
			results.ParsedFiles[fname].Error = fmt.Sprintf("unsupported filename/extension pattern for extraction from application/octet-stream: %s/%s", fileName, fileExtension)
			return
		}

	// Skip the files which aren't archives (or if they are archives, we have reasons for not wanting to unpack them)
	case "image/png",
		"image/jpeg",
		"image/jxr",
		"application/pdf",
		"application/vnd.ms-outlook", // We don't try to unpack Outlook .msg files
		"text/xml; charset=utf-8",
		"text/plain; charset=utf-8":
		results.ParsedFiles[fname].Supported = true
		return
	default:
		results.ParsedFiles[fname].Error = fmt.Sprintf("unsupported file type for extraction: %s", mType.String())
		return
	}
	err = unpackerImpl.UnpackStream(fname, inFile, fileSize, results, queue)
	return
}

// Use the reader to determine the type, then rewind the reader before returning. Only return
// error if the seeker can't be rewound. If the MIME type determinatin errors, return nil.
func getTypeFromReader(fname string, rdr io.ReadSeeker) (mType *mimetype.MIME, err error) {
	mType, err = mimetype.DetectReader(rdr)
	if err != nil {
		slog.Error("getting mime type", "file_name", fname, "error", err.Error())
		mType = nil
		err = nil
	}
	_, err = rdr.Seek(0, io.SeekStart)
	return
}

func copyFileToDir(fname, destDir string) (err error) {
	// Check and open the source file.
	srcStat, err := os.Stat(fname)
	if err != nil {
		return
	}
	if !srcStat.Mode().IsRegular() {
		err = fmt.Errorf("not a regular file: %s", fname)
		return
	}
	srcHandle, err := os.Open(fname)
	if err != nil {
		return
	}
	defer srcHandle.Close()

	// Create and open the destination file.
	destHandle, err := os.Create(path.Join(destDir, path.Base(fname)))
	if err != nil {
		return
	}
	defer destHandle.Close()

	// Do the copy.
	_, err = io.Copy(destHandle, srcHandle)
	return
}
