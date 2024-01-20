package parsers

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"log/slog"
	"path"
)

// OLE 1.0 reader. This implements io.Reader and looks a
// little like other archive readers WRT representation
// of metadata and accessing the enclosed object data.
type Ole10 struct {
	// Name of the enclosed file.
	Name string

	// Ostensibly, the temp path that this file was
	// loaded in from by the packaging program. Docs
	// are a little light on this part, so this could
	// mean something different between implementations.
	TempPath string

	// Ostensibly, the cache path on the original system.
	// Docs are a little light on this part, so this could
	// mean something different between implementations.
	CachePath string

	// The size of the enclosed object.
	Size int64

	// This holds a pointer to our OLE 1.0 file buffer.
	in *bytes.Buffer

	// This tracks our position reading the
	// enclosed object from the OLE container.
	readPosition int64
}

// Create a new OLE 1.0 reader.
//
// This constructor will attempt to parse the OLE 1.0 archive
// and perform metadata extraction before returning the new Ole10
// pointer. If the archive is corrupt or does not conform to the
// author's expectations and assumptions about how such a thing
// is structured, an error may be returned.
//
//	Args:
//		in (*bytes.Buffer):	This is the buffer containing OLE 1.0 file contents.
//
//	Returns:
//		o (*Ole10):		Reader for the OLE 1.0 file data.
//		err (error):	Failure to parse metadata from the file will cause this to be non-nil.
func NewOle10(in *bytes.Buffer) (o *Ole10, err error) {
	o = &Ole10{in: in}
	err = o.parse()
	return
}

// Read implements the io.Reader interface, taking care
// not to overrun the body of the enclosed object.
func (o *Ole10) Read(p []byte) (n int, err error) {
	lenRemains := o.Size - o.readPosition
	// If len(p) is 0, return 0/nil ()
	if len(p) == 0 {
		return
	}
	// If no embedded file content remains to be parsed,
	// return 0/io.EOF
	if lenRemains == 0 {
		n = 0
		err = io.EOF
		return
	}
	// If p asks for more bytes than remain in the embedded
	// file, we use a LimitReader to stop before overrunning
	// the payload.
	if len(p) > int(lenRemains) {
		r := io.LimitReader(o.in, lenRemains)
		n, err = r.Read(p)
		o.readPosition += int64(n)
		return
	}
	// If p does not ask for more bytes than remain,
	// just use Read.
	n, err = o.in.Read(p)
	o.readPosition += int64(n)
	return
}

// Parse and set the metadata from the OLE 1.0 object
// and advance the buffer to prepare for extraction.
func (o *Ole10) parse() (err error) {
	// Advance past the first header
	var firstHeader struct {
		Whatever [6]byte
	}
	err = binary.Read(o.in, binary.LittleEndian, &firstHeader)
	if err != nil {
		err = fmt.Errorf("%w: taking the first header", err)
		return
	}

	// The next section before null is the member name.
	name := bytes.NewBufferString("")
	err = advanceBufferToNextNull(o.in, name)
	if err != nil {
		err = fmt.Errorf("%w: Getting the member name", err)
		return
	}
	// Make sure that if any path elements are squeezed into this field,
	// we sanitize them out.
	o.Name = path.Base(name.String())

	// The next section is the original object path. This is sometimes interesting,
	// but not for this particular use case.
	cachePath := bytes.NewBufferString("")
	err = advanceBufferToNextNull(o.in, cachePath)
	if err != nil {
		err = fmt.Errorf("%w: getting the original path of the object", err)
		return
	}
	o.CachePath = cachePath.String()

	// Advance the stream a couple of null-terminated sections
	err = advanceBufferToNextNull(o.in, nil)
	if err != nil {
		err = fmt.Errorf("%w: skipping to null (1)", err)
		return
	}
	err = advanceBufferToNextNull(o.in, nil)
	if err != nil {
		err = fmt.Errorf("%w: skipping to null (2)", err)
		return
	}

	// We find the second 6-byte header. Get it and skip it.
	var secondHeader struct {
		Whatever [6]byte
	}
	err = binary.Read(o.in, binary.LittleEndian, &secondHeader)
	if err != nil {
		err = fmt.Errorf("%w: taking the second header", err)
		return
	}

	// The next null-terminated section is the shortened temp file name
	tempPath := bytes.NewBufferString("")
	err = advanceBufferToNextNull(o.in, tempPath)
	if err != nil {
		err = fmt.Errorf("%w: skipping to null (3 - shortened file name)", err)
		return
	}
	o.TempPath = tempPath.String()

	// Get the next section ahead of the object itself out of the way.
	var size uint32
	err = binary.Read(o.in, binary.LittleEndian, &size)
	if err != nil {
		err = fmt.Errorf("%w: getting the enclosed object size", err)
		slog.Error("getting object size", "error", err.Error())
		return
	}
	o.Size = int64(size)
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
