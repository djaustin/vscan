package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

type VLAN struct {
	ID, Name, Slug string
}

func slugify(s string) string {
	whiteSpace := regexp.MustCompile(`\W+`)
	slug := whiteSpace.ReplaceAllString(s, "-")
	return strings.ToLower(slug)
}

func NewVLAN(id, name string) VLAN {
	return VLAN{
		ID:   id,
		Name: name,
		Slug: slugify(name),
	}
}

func vlanDefinitionFound(row []string, column int, cell string) bool {
	vlanCellFound := strings.EqualFold("vlan", strings.TrimSpace(cell))
	descCellFound := strings.EqualFold("description", strings.TrimSpace(row[column+2]))
	vlanValueFound := strings.TrimSpace(row[column+1]) != ""
	descValueFound := strings.TrimSpace(row[column+3]) != ""

	return vlanCellFound && vlanValueFound && descCellFound && descValueFound

}

func scanRow(row []string) ([]VLAN, error) {
	var vlans []VLAN
	for i, cell := range row {
		if len(row) <= i+3 {
			continue
		}
		if vlanDefinitionFound(row, i, cell) {
			id, desc := row[i+1], row[i+3]
			vlan := NewVLAN(id, desc)
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
	if err := writer.Write([]string{"id", "name", "slug"}); err != nil {
		return err
	}
	defer writer.Flush()

	for _, vlan := range v {
		if err := writer.Write([]string{vlan.ID, vlan.Name, vlan.Slug}); err != nil {
			return err
		}
	}
	return nil
}

func (v VLANSet) PrintTable() {
	fmt.Printf("|%30s|%30s|%30s|\n", "VLAN", "Name", "Slug")
	fmt.Printf("|%s|\n", strings.Repeat("-", 92))
	for _, vlan := range v {
		fmt.Printf("|%30s|%30s|%30s|\n", vlan.ID, vlan.Name, vlan.Slug)
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
