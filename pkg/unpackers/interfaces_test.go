package unpackers

import "testing"

// Make sure all of our implementations conform.
func TestInterfaceImplementations(t *testing.T) {
	// Check OLE 1.0 unpacker
	var _ Unpacker = (*OLE10Native)(nil)

	// Check MSCFB unpacker
	var _ Unpacker = (*MSCFB)(nil)

	// Check OfficeZip unpacker
	var _ Unpacker = (*OfficeZip)(nil)
}
