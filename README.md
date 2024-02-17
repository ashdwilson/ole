# ole

A tool for unpacking Office files.

## Motivation

During the course of my work at @UrbaneSec, I have needed to parse a variety of files from a variety of sources.

Sometimes we need to unpack MS Office files. Maybe it's because we're simply curious what's inside. Maybe certain embedded objects just don't unpack because we aren't running a software stack with OLEv1.0 support. Office documents have proven to be really useful containers for conveying all sorts of information. Screenshots. PDFs. So many different kinds of files can be dragged and dropped into a Word doc and emailed out into the world. There's a problem, though: Office doesn't have absolute feature parity across all supported platforms. Extracting some object types from a given Office document can be difficult if you're running Office on a Mac.

This tool will extract all members from all recognized archive formats, recursively expanding everything in a given MS Office document. It's not pretty (creates a lot of files), but it does the thing it's supposed to do and it's easy enough to extend.

The code in this repository shold be treated as experimental. While some basic testing is in-place and some care has been taken to make sure output paths are sane, this has not undergone a rigorous audit with maliciously-formed content in mind. Use at your own risk.

## Installation

### Prerequisites to

This tool was developed on MacOS. It may work on other platforms, and probably works fine on Linux, given the tests are configured to run on Ubuntu.

Go 1.21 or newer is required.

### The act of

`go install github.com/ashdwilson/ole@main`

## Usage

```
Unpack an Office document, including any embedded OLE objects.

Usage:
  ole [flags]

Flags:
  -h, --help            help for ole
  -i, --infile string   Input Office format file.
  -o, --outdir string   Output directory for extracted assets.
```

## TODO

- [ ] Capture trailing data (OLE v1)
- [ ] Better labeling for extracted metadata (OLE v1)