//go:generate ../../bin/goversioninfo.exe -icon=image.ico -manifest=referentiel-sncf.exe.manifest -o=referentiel-sncf.syso

package main

import (
	"fmt"
	"html/template"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"referentiel-sncf/icon"
	"runtime"
	"strings"
	"time"

	"github.com/cratonica/trayhost"
	"github.com/tealeg/xlsx"
	"github.com/toqueteos/webbrowser"
)

type Tableur struct {
	Reference string
	IndexUser string
	Title     string
	Version   string
	DateApp   string
	Tiroir    string
	Classeur  string
	Link      string
}

const (
	OneDayInSecond      = 86400
	FirstDayUnixInExcel = 25569
)

func main() {

	runtime.LockOSThread()
	port := ":1981"

	http.HandleFunc("/", HomeHandler)

	go http.ListenAndServe(port, nil)
	//Trayhost Icon dans barre des tâches
	go func() {
		trayhost.SetUrl("http://localhost" + port)
	}()
	webbrowser.Open("http://localhost" + port)
	trayhost.EnterLoop("Referentiel SNCF", icon.IconData)
	fmt.Println("Exiting")

}

func ReadInXlsx() []*Tableur {
	excelFileName := "./srv.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Fatal(err)
	}
	tableur := []*Tableur{}

	for _, sheet := range xlFile.Sheets {
		for r, row := range sheet.Rows {
			if r > 0 {
				tab := new(Tableur)
				tab.Reference = row.Cells[0].Value
				tab.IndexUser = row.Cells[1].Value
				tab.Title = strings.Title(strings.ToLower(row.Cells[2].Value))
				tab.Version = row.Cells[3].Value
				if row.Cells[4].GetNumberFormat() == "mm-dd-yy" {
					day, _ := row.Cells[4].Int64()
					tab.DateApp = transformDateXlsToFormat(day)
				} else {
					tab.DateApp = row.Cells[4].Value
				}
				tab.Tiroir = row.Cells[5].Value
				tab.Classeur = row.Cells[6].Value
				s := strings.Split(row.Cells[7].Formula(), "\"")
				for k, _ := range s {
					if k == 1 {
						tab.Link = s[k]
					}
				}
				tableur = append(tableur, tab)
			}
		}
	}
	return tableur
}

func transformDateXlsToFormat(i int64) string {
	d := (i - FirstDayUnixInExcel) * OneDayInSecond
	date := time.Unix(d, 0)
	return date.Format("02-01-2006")
}

func HomeHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	t, err := template.New("page").Parse(PAGE)
	if err != nil {
		log.Println(err)
	}

	tab := ReadInXlsx()

	model := map[string]interface{}{
		"Title": "Référentiel SNCF",
		"Tab":   tab,
	}
	err = t.Execute(w, model)

	Check("Home : ", err)
}

func Check(function string, err error) error {
	logFile := filepath.Clean("./logFileError.log")
	if _, e := os.Stat(logFile); os.IsNotExist(e) {
		w, e := os.Create(logFile)
		Check("check : ", e)
		defer w.Close()
	}

	f, e := os.OpenFile(logFile, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0666)
	if e != nil {
		Check("check : ", e)
	}
	if err != nil {
		log.SetOutput(f)
		log.Print(function, err, "\n\r")
		return err
	}
	defer f.Close()
	return nil
}
