package main

import (
	"archive/zip"
	"flag"
	"fmt"
	"io"
	"regexp"
)

func findVer(in string) string {
	regApp := regexp.MustCompile(`<Application>(.+?)</Application>`)
	regVer := regexp.MustCompile(`<AppVersion>(.+?)</AppVersion>`)
	app := regApp.FindAllStringSubmatch(in, 1)
	ver := regVer.FindAllStringSubmatch(in, 1)

	return fmt.Sprintf("%s - %s", app[0][1], ver[0][1])
}

func main() {
	var Ifile string
	var Ihelp bool

	flag.StringVar(&Ifile, "file", "./test.docx", "path to file")
	flag.BoolVar(&Ihelp, "help", false, "this help")
	flag.Parse()

	if Ihelp {
		flag.Usage()
	}

	r, err := zip.OpenReader(Ifile)
	if err != nil {
		fmt.Printf("Cann't open input file '%s': %s", Ifile, err)
		return
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "docProps/app.xml" {
			rc, err := f.Open()
			if err != nil {
				fmt.Printf("Error when open file '%s' in zip", f.Name)
				return
			}
			defer r.Close()

			bData, err := io.ReadAll(rc)
			if err != nil {
				fmt.Printf("Error read zip file '%s'", f.Name)
				return
			}
			sData := string(bData)
			fmt.Println(findVer(sData))
		}
	}
}
