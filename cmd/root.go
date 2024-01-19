/*
Copyright © 2024 Ash Wilson <@ashdwilson>
*/
package cmd

import (
	"os"

	"github.com/ashdwilson/ole/internal/unpacker"
	"github.com/spf13/cobra"
)

var inFile, outDir string

// rootCmd represents the base command when called without any subcommands
var rootCmd = &cobra.Command{
	Use:   "ole",
	Short: "Unpack a MS office document, including OLE objects.",
	RunE:  unpack,
}

func Execute() {
	err := rootCmd.Execute()
	if err != nil {
		os.Exit(1)
	}
}

func init() {
	rootCmd.Flags().StringVarP(&inFile, "infile", "i", "", "Input MS office format file.")
	rootCmd.Flags().StringVarP(&outDir, "outdir", "o", "", "Output directory for extracted assets.")
}

func unpack(cmd *cobra.Command, args []string) (err error) {
	err = unpacker.Unpack(inFile, outDir)
	return
}
