package unpackers

import (
	"container/list"
	"io"

	"github.com/ashdwilson/ole/pkg/models"
)

type Unpacker interface {
	UnpackStream(inpath string, stream io.ReaderAt, size int64, results *models.Results, queue *list.List) (err error)
}
