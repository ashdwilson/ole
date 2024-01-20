package unpackers

import (
	"bytes"
	"container/list"
	"errors"
	"fmt"
	"io"
	"os"
	"path"

	"github.com/ashdwilson/ole/pkg/models"
	"github.com/ashdwilson/ole/pkg/parsers"
)

type OLE10Native struct{}

func (o *OLE10Native) UnpackStream(inpath string, stream io.ReaderAt, size int64, results *models.Results, queue *list.List) (err error) {
	// Set a base path for extracting embedded objects.
	basePath := fmt.Sprintf("%s-members", inpath)
	err = os.MkdirAll(basePath, 0770)
	if err != nil {
		return
	}
	errs := []error{}

	// Turn the io.ReaderAt into a buffer.
	streamBuf := make([]byte, size)
	_, err = stream.ReadAt(streamBuf, 0)
	if err != nil {
		return
	}
	s := bytes.NewBuffer(streamBuf)

	oleReader, err := parsers.NewOle10(s)
	if err != nil {
		err = fmt.Errorf("%w: parsing OLE 1.0 container", err)
		return
	}

	// Create an output file and read the object into it
	fullNewFilePath := path.Join(basePath, oleReader.Name)
	outFile, err := os.Create(fullNewFilePath)
	if err != nil {
		err = fmt.Errorf("%w: creating a new file to receive the object", err)
		return
	}
	_, err = io.Copy(outFile, oleReader)
	if err != nil {
		err = fmt.Errorf("%w: writing the extracted object to a file", err)
		return
	}

	queue.PushBack(fullNewFilePath)

	results.ParsedFiles[inpath].Expanded = true
	err = errors.Join(errs...)
	return
}
