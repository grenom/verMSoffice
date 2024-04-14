package main

import (
	"archive/zip"
	"flag"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

func getInfo(Ifile string) (string, error) {
	var out string

	r, err := zip.OpenReader(Ifile)
	if err != nil {
		fmt.Printf("Cann't open input file '%s': %s", Ifile, err)
		return "", err
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "docProps/app.xml" {
			rc, err := f.Open()
			if err != nil {
				fmt.Printf("Error when open file '%s' in zip", f.Name)
				return "", err
			}
			defer r.Close()

			bData, err := io.ReadAll(rc)
			if err != nil {
				fmt.Printf("Error read zip file '%s'", f.Name)
				return "", err
			}
			sData := string(bData)
			//out = findVer(sData)
			out = findEditor(sData)
		}
	}
	return out, nil
}

func findVer(in string) string {
	regApp := regexp.MustCompile(`<Application>(.+?)</Application>`)
	regVer := regexp.MustCompile(`<AppVersion>(.+?)</AppVersion>`)
	app := regApp.FindAllStringSubmatch(in, 1)
	ver := regVer.FindAllStringSubmatch(in, 1)

	return fmt.Sprintf("%s - %s", app[0][1], ver[0][1])
}

func findEditor(in string) string {
	regApp := regexp.MustCompile(`(?i)worksheets`)
	found := regApp.MatchString(in)
	if found {
		return "P7 office"
	} else {
		return "MS office"
	}
}

func main() {
	var target_files []string
	var Ihelp bool

	flag.BoolVar(&Ihelp, "help", false, "this help")
	flag.Parse()

	if Ihelp {
		fmt.Println("This application searches an editor in Office files")
		flag.Usage()
	}

	files, err := os.ReadDir("./")
	if err != nil {
		fmt.Printf("Error list files: %s", err)
		os.Exit(1)
	}
	for _, file := range files {
		if file.Type().IsRegular() && (strings.HasSuffix(file.Name(), ".xlsx") || strings.HasSuffix(file.Name(), ".docx")) {
			target_files = append(target_files, file.Name())
		}
	}
	if len(target_files) == 0 {
		fmt.Println("Not found office files there.")
	}

	for _, file := range target_files {
		info, err := getInfo(file)
		if err != nil {
			fmt.Printf("Error process '%s' file: %s", file, err)
		}
		fmt.Printf("%s - %s\n", file, info)
	}

	var tmp_exit string
	fmt.Scanln(&tmp_exit)
}
