package parsers

import (
	"bytes"
	"crypto/sha256"
	"fmt"
	"io"
	"os"
	"path"
	"testing"
)

var pathToSampleDataDir = "../../test/data/"

// A simple test, which verifies that the unpacked
// asset carries the expected metadata and payload.
func TestOle10E2E(t *testing.T) {
	expectedName := "Untitled.msg"
	expectedTempPath := `C:\Users\ADMINI~1\AppData\Local\Temp\2\{E4DFC792-E2A9-41CE-A805-A6653662797E}\{E8FC9E50-3E3D-414E-B742-E28A8ED01A62}\Untitled.msg`
	expectedCachePath := `C:\Users\Administrator\AppData\Local\Microsoft\Windows\INetCache\Content.Word\Untitled.msg`
	controlFile := path.Join(pathToSampleDataDir, "sample1.msg")
	cf, err := os.Open(controlFile)
	if err != nil {
		t.Fatal(err)
	}
	defer cf.Close()
	cfc, err := io.ReadAll(cf)
	if err != nil {
		t.Fatal(err)
	}
	infile := path.Join(pathToSampleDataDir, "sample1.ole")
	f, err := os.Open(infile)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	fc, err := io.ReadAll(f)
	if err != nil {
		t.Fatal(err)
	}
	b := bytes.NewBuffer(fc)
	o, err := NewOle10(b)
	if err != nil {
		t.Fatal(err)
	}
	if o.Name != expectedName {
		t.Errorf("name mismatch - expected %s got %s", expectedName, o.Name)
	}
	if o.CachePath != expectedCachePath {
		t.Errorf("cache path mismatch - expected %s got %s", expectedCachePath, o.CachePath)
	}
	if o.TempPath != expectedTempPath {
		t.Errorf("temp path mismatch - expected %s got %s", expectedTempPath, o.TempPath)
	}
	if o.Size == 0 {
		t.Errorf("size not being set")
	}
	embedded, err := io.ReadAll(o)
	if err != nil {
		t.Fatal(err)
	}
	h := sha256.New()
	h.Write(embedded)
	embeddedSum := fmt.Sprintf("%x", h.Sum(nil))
	ch := sha256.New()
	ch.Write(cfc)
	controlSum := fmt.Sprintf("%x", ch.Sum(nil))

	if embeddedSum != controlSum {
		t.Errorf("hash mismatch - expected %s got %s", controlSum, embeddedSum)
	}
	if len(embedded) != int(o.Size) {
		t.Errorf("mismatch between header-indicated size and actual size - got %d expected %d", len(embedded), o.Size)
	}
}
