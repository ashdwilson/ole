package unpackers

import (
	"container/list"
	"errors"
	"fmt"
	"io"
	"os"
	"path"

	"github.com/ashdwilson/ole/pkg/models"
	"github.com/richardlehane/mscfb"
)

// The MSCFB implementation of Unpacker uses a 3rd-party library
// to parse MS-CFB (OLE v2) files.
type MSCFB struct{}

// This unpacker extracts all enclosed objects, and enqueues them for further examination.
func (m *MSCFB) UnpackStream(inpath string, stream io.ReaderAt, size int64, results *models.Results, queue *list.List) (err error) {
	// Base path is inpath-members/
	basePath := fmt.Sprintf("%s-members", inpath)
	err = os.MkdirAll(basePath, 0770)
	if err != nil {
		return
	}
	errs := []error{}

	// Open as a zip archive
	var rdr *mscfb.Reader
	rdr, err = mscfb.New(stream)
	if err != nil {
		return
	}

	// Iterate through members
	for entry, err := rdr.Next(); err == nil; entry, err = rdr.Next() {

		if entry.FileInfo().IsDir() {
			// Create a dir
			dName := path.Join(basePath, entry.Name)
			err = os.MkdirAll(dName, 0770)
			if err != nil {
				err = fmt.Errorf("%w: creating directory %s", err, dName)
				errs = append(errs, err)
			}
			continue
		}
		pathElements := []string{basePath}
		pathElements = append(pathElements, entry.Path...)
		pathElements = append(pathElements, entry.Name)
		newFilePath := path.Join(pathElements...)

		var newFile *os.File
		newFile, err = os.Create(newFilePath)
		if err != nil {
			err = fmt.Errorf("%w: creating file %s", err, newFilePath)
			errs = append(errs, err)
			continue
		}

		_, err = io.Copy(newFile, entry)
		if err != nil {
			err = fmt.Errorf("%w: extracting member %s to %s", err, entry.Name, newFilePath)
			errs = append(errs, err)
			continue
		}
		queue.PushBack(newFilePath)
	}
	results.ParsedFiles[inpath].Expanded = true
	err = errors.Join(errs...)
	return
}
