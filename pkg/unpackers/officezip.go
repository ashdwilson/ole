package unpackers

import (
	"archive/zip"
	"container/list"
	"errors"
	"fmt"
	"io"
	"os"
	"path"

	"github.com/ashdwilson/ole/pkg/models"
)

// This implementation of the Unpacker interface uses archive/zip
// to extract all members from the Office document.
type OfficeZip struct{}

// Unpack all the archive members and queue them up for parsing.
func (o *OfficeZip) UnpackStream(inpath string, stream io.ReaderAt, size int64, results *models.Results, queue *list.List) (err error) {
	// Base path is inpath-members/
	basePath := fmt.Sprintf("%s-members", inpath)
	err = os.MkdirAll(basePath, 0770)
	if err != nil {
		return
	}
	errs := []error{}

	// Open as a zip archive
	var rdr *zip.Reader
	rdr, err = zip.NewReader(stream, size)

	// Iterate through members
	for _, f := range rdr.File {
		rstat := f.FileInfo()
		if rstat.IsDir() {
			dName := path.Join(basePath, rstat.Name())
			err = os.MkdirAll(dName, 0770)
			if err != nil {
				err = fmt.Errorf("%w: creating directory %s", err, dName)
				errs = append(errs, err)
			}
			continue
		}
		newFilePath := path.Join(basePath, rstat.Name())
		var newFile *os.File
		newFile, err = os.Create(newFilePath)
		if err != nil {
			err = fmt.Errorf("%w: creating file %s", err, newFilePath)
			errs = append(errs, err)
			continue
		}
		fHandle, err := f.Open()
		if err != nil {
			err = fmt.Errorf("%w: opening archive member %s", err, f.Name)
			errs = append(errs, err)
			continue
		}
		defer fHandle.Close()
		_, err = io.Copy(newFile, fHandle)
		if err != nil {
			err = fmt.Errorf("%w: extracting member %s to %s", err, f.Name, newFilePath)
			errs = append(errs, err)
			continue
		}
		queue.PushBack(newFilePath)
	}
	results.ParsedFiles[inpath].Expanded = true
	err = errors.Join(errs...)
	return
}
