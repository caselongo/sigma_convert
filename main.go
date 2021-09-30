package main

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"math"
	"os"
	"reflect"
	"strconv"
	"time"
	"unicode/utf8"

	excelize "github.com/360EntSecGroup-Skylar/excelize/v2"
)

const (
//fileName string = "C:/Users/michi/Downloads/2021_08_16__10_41.tcx"
//density int = 2
)

type TrainingCenterDatabase struct {
	TrainingCenterDatabase xml.Name   `xml:"TrainingCenterDatabase"`
	Activities             Activities `xml:"Activities"`
}

type Activities struct {
	XMLName  xml.Name   `xml:"Activities"`
	Activity []Activity `xml:"Activity"`
}

type Activity struct {
	XMLName xml.Name `xml:"Activity"`
	ID      string   `xml:"Id"`
	Lap     []Lap    `xml:"Lap"`
}

type Lap struct {
	XMLName          xml.Name `xml:"Lap"`
	StartTime        string   `xml:"StartTime,attr"`
	Track            Track    `xml:"Track"`
	TotalTimeSeconds float64  `xml:"TotalTimeSeconds"`
	DistanceMeters   float64  `xml:"DistanceMeters"`
}

type Track struct {
	XMLName    xml.Name     `xml:"Track"`
	Trackpoint []Trackpoint `xml:"Trackpoint"`
}

type Trackpoint struct {
	XMLName  xml.Name `xml:"Trackpoint"`
	Time     string   `xml:"Time"`
	Position struct {
		XMLName          xml.Name `xml:"Position"`
		LatitudeDegrees  float64  `xml:"LatitudeDegrees"`
		LongitudeDegrees float64  `xml:"LongitudeDegrees"`
	} `xml:"Position"`
	AltitudeMeters float64 `xml:"AltitudeMeters"`
	DistanceMeters float64 `xml:"DistanceMeters"`
}

type Row struct {
	Latitude  float64
	Longitude float64
	distance  float64
	Distance  float64
	Altitude  float64
	Time      int64
	Marker    string
}

func main() {
	if len(os.Args) != 4 {
		fmt.Println("Invalid number of arguments. Required: input file, output file, density")
		return
	}

	fileName := os.Args[1]
	outputFileName := os.Args[2]
	_density, err := strconv.ParseInt(os.Args[3], 10, 64)
	if err != nil {
		fmt.Println(err)
		return
	}
	density := int(_density)

	// Open the xmlFile
	xmlFile, err := os.Open(fileName)

	// if os.Open returns an error then handle it
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("\tSuccessfully opened xml")
	// defer the closing of xmlFile so that we can parse it.
	defer xmlFile.Close()

	byteValue, _ := ioutil.ReadAll(xmlFile)

	// Unmarshal takes a []byte and fills the rss struct with the values found in the xmlFile
	trainingCenterDatabase := TrainingCenterDatabase{}

	decoder := xml.NewDecoder(bytes.NewReader(byteValue))
	err = decoder.Decode(&trainingCenterDatabase)
	if err != nil {
		fmt.Println(err)
		return
	}

	rows := []Row{}
	var prevRow *Row = nil
	rowNumber := 0

	for _, activity := range trainingCenterDatabase.Activities.Activity {
		for _, lap := range activity.Lap {
			for i, trackpoint := range lap.Track.Trackpoint {
				rowNumber++

				t, _ := time.Parse("2006-01-02T15:04:05Z", trackpoint.Time)
				marker := ""
				if rowNumber == 1 {
					marker = "x"
				}

				row := Row{
					Latitude:  trackpoint.Position.LatitudeDegrees,
					Longitude: trackpoint.Position.LongitudeDegrees,
					distance:  trackpoint.DistanceMeters,
					Distance:  math.Round(trackpoint.DistanceMeters/10) / 100,
					Altitude:  trackpoint.AltitudeMeters,
					Time:      t.Unix(),
					Marker:    marker,
				}

				if i == 0 && prevRow != nil {
					t, err := time.Parse("2006-01-02T15:04:05Z", lap.StartTime)
					if err != nil {
						fmt.Println(err)
					} else {
						dt := row.Time - prevRow.Time
						d := float64(t.Unix()-prevRow.Time) / float64(dt)
						dPrev := float64(row.Time-t.Unix()) / float64(dt)

						rows = append(rows, Row{
							Latitude:  row.Latitude*d + prevRow.Latitude*dPrev,
							Longitude: row.Longitude*d + prevRow.Longitude*dPrev,
							Distance:  math.Round((row.distance*d+prevRow.distance*dPrev)/10) / 100,
							Altitude:  math.Round(row.Altitude*d + prevRow.Altitude*dPrev),
							Time:      t.Unix(),
							Marker:    "x",
						})
					}
				}

				if density > 1 {
					if (rowNumber-1)%density != 0 {
						continue
					}
				}

				rows = append(rows, row)

				prevRow = &row

			}
		}
	}

	CreateExcelFile(&rows, "data", outputFileName)
}

func CreateExcelFile(data interface{}, sheetName string, fileName string) {
	structSlice := reflect.ValueOf(data).Elem()
	structType := reflect.TypeOf(data).Elem().Elem()

	excelFile := excelize.NewFile()
	// Create a new sheet.
	excelFile.NewSheet(sheetName)
	// delete all other sheets
	sheets := excelFile.GetSheetList()
	for _, sheet := range sheets {
		if sheet == sheetName {
			continue
		}

		excelFile.DeleteSheet(sheet)
	}

	rowIndex := 1
	columnIndex := 0
	for i := 0; i < structType.NumField(); i++ {
		fieldName := structType.Field(i).Tag.Get("xlsx")
		if fieldName == "" {
			fieldName = structType.Field(i).Name
		}

		structValue := structSlice.Index(i)
		field := reflect.ValueOf(structValue.Addr().Interface()).Elem().FieldByIndex([]int{i})
		if !field.CanSet() {
			continue
		}

		columnName, _ := excelize.ColumnNumberToName(columnIndex + 1)
		cellName := fmt.Sprintf("%s%v", columnName, rowIndex)

		err := excelFile.SetCellValue(sheetName, cellName, fieldName)
		if err != nil {
			fmt.Println(err)
			return
		}
		columnIndex++
	}

	rowIndex++

	for i := 0; i < structSlice.Len(); i++ {
		structValue := structSlice.Index(i)

		columnIndex := 0
		for j := 0; j < structValue.NumField(); j++ {
			field := reflect.ValueOf(structValue.Addr().Interface()).Elem().FieldByIndex([]int{j})
			if !field.CanSet() {
				continue
			}

			columnName, _ := excelize.ColumnNumberToName(columnIndex + 1)
			cellName := fmt.Sprintf("%s%v", columnName, rowIndex)

			columnIndex++

			if field.IsZero() {
				continue
			}

			err := excelFile.SetCellValue(sheetName, cellName, field.Interface())
			if err != nil {
				fmt.Println(err)
				return
			}
		}

		rowIndex++
	}

	// Autofit all columns according to their text content
	cols, err := excelFile.GetCols(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	for idx, col := range cols {
		largestWidth := 0
		for _, rowCell := range col {
			cellWidth := utf8.RuneCountInString(rowCell) + 2 // + 2 for margin
			if cellWidth > largestWidth {
				largestWidth = cellWidth
			}
		}
		name, err := excelize.ColumnNumberToName(idx + 1)
		if err != nil {
			fmt.Println(err)
			return
		}
		excelFile.SetColWidth(sheetName, name, name, float64(largestWidth))
	}

	// Convert excel file to bytes
	buffer, err := excelFile.WriteToBuffer()
	if err != nil {
		fmt.Println(err)
		return
	}
	b, err := ioutil.ReadAll(buffer)
	if err != nil {
		fmt.Println(err)
		return
	}

	// write to file
	err = ioutil.WriteFile(fileName, b, 0644)
	if err != nil {
		fmt.Println(err)
	}
}
