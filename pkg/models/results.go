package models

// Scan Results
type Results struct {
	ParsedFiles map[string]*Result
}

// A single result item
type Result struct {
	// MIME type of file
	FileType string

	// Was the file expanded successfully?
	Expanded bool

	// Is extraction supported?
	Supported bool

	// Any error output
	Error string
}
