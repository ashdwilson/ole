package unpackers

import (
	"container/list"
	"io"

	"github.com/ashdwilson/ole/pkg/models"
)

// An Unpacker is a component that reads an archive and extracts
// the enclosed objects.
//
// This interface describes how the unpackers in this software project should work.
// Generally, an unpacker works like this:
//   - Create a new directory for all member files to be extracted to (${inpath}-members/)
//   - Load the abstraction specific to the data stream we are trying to unpack (archive/zip, for instance)
//   - For each enclosed file, write the contents to a new file under the ${inpath}-members/ directory.
//   - The naming of the file is left up to the implementation, and should take care to avoid collisions.
type Unpacker interface {
	// UnpackStream extracts members from the archive stream:
	//
	//	Args:
	//		inpath (string):		Path to original file on disk. Used to construct the output directory.
	//		stream (io.ReaderAt):	This is the data stream we will extract archive members from.
	//		size (int64):		This is the total size of the archive file conveyed over `stream`.
	//		results (*models.Results)	This is the object we use to track overall results of the extraction.
	//		queue (*list.List)		The implementation adds extracted files to the queue for recursive processing.
	//
	//	Returns:
	//		err (error)
	UnpackStream(inpath string, stream io.ReaderAt, size int64, results *models.Results, queue *list.List) (err error)
}
