package unpackers

import (
	"bytes"
	"container/list"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"log/slog"
	"os"
	"path"

	"github.com/ashdwilson/ole/pkg/models"
)

type OLE10Native struct{}

func (o *OLE10Native) UnpackStream(inpath string, stream io.ReaderAt, size int64, results *models.Results, queue *list.List) (err error) {
	// Base path is inpath-members/
	basePath := fmt.Sprintf("%s-members", inpath)
	err = os.MkdirAll(basePath, 0770)
	if err != nil {
		return
	}
	errs := []error{}

	streamBuf := make([]byte, size)
	_, err = stream.ReadAt(streamBuf, 0)
	if err != nil {
		return
	}
	s := bytes.NewBuffer(streamBuf)
	var firstHeader struct {
		Whatever [6]byte
	}
	var secondHeader struct {
		Whatever [6]byte
	}

	// Take the first header
	err = binary.Read(s, binary.LittleEndian, &firstHeader)
	if err != nil {
		err = fmt.Errorf("%w: taking the first header", err)
		return
	}

	// The next section before null is the member name.
	name := bytes.NewBufferString("")
	err = advanceBufferToNextNull(s, name)
	if err != nil {
		err = fmt.Errorf("%w: Getting the member name", err)
		return
	}
	objectName := name.String()

	// Create a path that represents the original object path.
	originalPath := bytes.NewBufferString("")
	err = advanceBufferToNextNull(s, originalPath)
	if err != nil {
		err = fmt.Errorf("%w: getting the original path of the object", err)
		return
	}

	// Advance the stream a couple of null-terminated sections
	err = advanceBufferToNextNull(s, nil)
	if err != nil {
		err = fmt.Errorf("%w: skipping to null (1)", err)
		return
	}
	err = advanceBufferToNextNull(s, nil)
	if err != nil {
		err = fmt.Errorf("%w: skipping to null (2)", err)
		return
	}

	// We find the second 6-byte header. Get it and skip it.
	err = binary.Read(s, binary.LittleEndian, &secondHeader)
	if err != nil {
		err = fmt.Errorf("%w: taking the second header", err)
		return
	}

	// The next null-terminated section is the shortened file name
	// We skip it.
	err = advanceBufferToNextNull(s, nil)
	if err != nil {
		err = fmt.Errorf("%w: skipping to null (3 - shortened file name)", err)
		return
	}

	// Get the next section ahead of the object itself out of the way.
	var sg struct {
		Whatever [4]byte
	}
	err = binary.Read(s, binary.LittleEndian, &sg)
	if err != nil {
		err = fmt.Errorf("%w: getting the enclosed object size", err)
		slog.Error("poops", "error", err.Error())
		return
	}

	// Create an output file and read the object into it
	fullNewFilePath := path.Join(basePath, objectName)
	outFile, err := os.Create(fullNewFilePath)
	if err != nil {
		err = fmt.Errorf("%w: Creating a new file to receive the object", err)
		return
	}
	_, err = io.Copy(outFile, s)
	if err != nil {
		err = fmt.Errorf("%w: writing the extracted object to a file", err)
		return
	}

	queue.PushBack(fullNewFilePath)

	results.ParsedFiles[inpath].Expanded = true
	err = errors.Join(errs...)
	return
}

// Advance the buffer to the next null. If w is provided, it will receive the
// contents of the buffer read until null was reached.
func advanceBufferToNextNull(r, w *bytes.Buffer) (err error) {
	result, err := r.ReadBytes('\x00')
	if err != nil {
		return
	}
	if w != nil {
		_, err = w.Write(result[:len(result)-1])
	}
	return
}
