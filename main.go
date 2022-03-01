package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

type VLAN struct {
	ID, Name string
}

func isVLANDefinition(row []string, column int, cell string) bool {
	return strings.EqualFold("vlan", cell) && strings.EqualFold("description", row[column+2])
}

func scanRow(row []string) ([]VLAN, error) {
	var vlans []VLAN
	for i, cell := range row {
		if len(row) < i+3 {
			continue
		}
		if isVLANDefinition(row, i, cell) {
			id, desc := row[i+1], row[i+3]
			vlan := VLAN{
				ID:   id,
				Name: desc,
			}
			vlans = append(vlans, vlan)
		}
	}
	return vlans, nil
}

func findVLANs(path string) ([]VLAN, error) {
	var vlans []VLAN
	file, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open file for parsing: %w", err)
	}
	sheets := file.GetSheetList()
	for _, v := range sheets {
		rows, err := file.GetRows(v)
		if err != nil {
			fmt.Print(err)
			continue
		}
		for _, row := range rows {
			detected, err := scanRow(row)
			if err != nil {
				fmt.Print(err)
				continue
			}
			vlans = append(vlans, detected...)
		}
	}
	return vlans, nil
}

type VLANSet []VLAN

func (v VLANSet) WriteCSV(path string) error {
	file, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	if err := writer.Write([]string{"id", "name"}); err != nil {
		return err
	}
	defer writer.Flush()

	for _, vlan := range v {
		if err := writer.Write([]string{vlan.ID, vlan.Name}); err != nil {
			return err
		}
	}
	return nil
}

func (v VLANSet) PrintTable() {
	fmt.Printf("|%30s|%30s|\n", "VLAN", "Description")
	fmt.Printf("|%s|\n", strings.Repeat("-", 61))
	for _, vlan := range v {
		fmt.Printf("|%30s|%30s|\n", vlan.ID, vlan.Name)
	}
}

func main() {
	inPath := flag.String("in", "", "path to the spreadsheet file to parse")
	outPath := flag.String("out", "", "path to the output file")
	flag.Parse()
	if *inPath == "" || *outPath == "" {
		flag.Usage()
		return
	}
	vlans, err := findVLANs(*inPath)
	if err != nil {
		fmt.Println(err)
	}
	vlanSet := VLANSet(vlans)
	vlanSet.PrintTable()
	err = vlanSet.WriteCSV(*outPath)
	if err != nil {
		fmt.Println(err)
	}
}
