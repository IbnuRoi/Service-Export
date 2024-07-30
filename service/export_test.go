package service

import (
	"fmt"
	"sam/RnD/codebase/service-export/models"
	"testing"

	"github.com/brianvoe/gofakeit/v7"
	"github.com/xuri/excelize/v2"
)

func TestExportLimit(t *testing.T) {
	totalUsers := 100
	var users []*models.User
	headers := []interface{}{"Number", "Name", "Email", "Gender", "Phone", "City", "State", "Country"}
	for i := 0; i < totalUsers; i++ {
		user := models.User{
			Name:   gofakeit.Name(),
			Email:  gofakeit.Email(),
			Gender: models.Gender(gofakeit.Gender()),
			Phone:  gofakeit.Phone(),
			Address: models.Address{
				City:    gofakeit.City(),
				State:   gofakeit.State(),
				Country: gofakeit.Country(),
			},
		}
		users = append(users, &user)
	}
	// Membuat instance baru Excel file
	xlsx := excelize.NewFile()

	// Menambahkan data ke file Excel
	for i, u := range users {
		// Menentukan nama sheet berdasarkan nomor urutan
		sheetIndex := i / 10
		sheetName := fmt.Sprintf("Sheet%d", sheetIndex)

		// Jika sheet belum ada, buat sheet baru
		if !sheetExists(xlsx, sheetName) {
			xlsx.NewSheet(sheetName)
		}
		writer, err := xlsx.NewStreamWriter(sheetName)
		if err != nil {
			panic("Gada sheet 2!")
		}
		setHeader(writer, headers)

		// Menambahkan data ke cell pada sheet yang sesuai
		// cell := fmt.Sprintf("%s%d", "A", i+2)
		value := []interface{}{i, u.Name, u.Email, u.Gender, u.Phone, u.Address.City, u.Address.State, u.Address.Country}
		// xlsx.SetCellValue(sheetName, cell, value)

		// Insert data
		streamValues(writer, i, 20, value)

		if err := writer.Flush(); err != nil {
			panic("Kaga ke flush gess!")
		}
	}

	// Simpan file Excel
	if err := xlsx.SaveAs("./output.xlsx"); err != nil {
		fmt.Println("Error:", err)
		return
	}

	fmt.Println("Excel file created successfully.")
}

func TestExport1(t *testing.T) {
	totalUsers := 100
	var users []*models.User
	headers := []interface{}{"Number", "Name", "Email", "Gender", "Phone", "City", "State", "Country"}
	for i := 0; i < totalUsers; i++ {
		user := models.User{
			Name:   gofakeit.Name(),
			Email:  gofakeit.Email(),
			Gender: models.Gender(gofakeit.Gender()),
			Phone:  gofakeit.Phone(),
			Address: models.Address{
				City:    gofakeit.City(),
				State:   gofakeit.State(),
				Country: gofakeit.Country(),
			},
		}
		users = append(users, &user)
	}
	fmt.Println(len(users))

	f := excelize.NewFile()

	valuePerSheet := 10
	numData := len(users)
	numSheet := numData / valuePerSheet

	for i := 0; i < numSheet-1; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i+1)
		writer, err := f.NewStreamWriter(sheetName)
		if err != nil {
			panic(err)
		}
		defer writer.Flush()

		setHeader(writer, headers)

		for row, user := range users {
			rowValues := []interface{}{row + 1, user.Name, user.Email, user.Gender, user.Phone, user.Address.City, user.Address.State, user.Address.Country}
			streamValues(writer, row, 20, rowValues)
		}
	}

	if err := f.SaveAs("./coba1.xlsx"); err != nil {
		panic(err)
	}

}

func TestExportCoba(t *testing.T) {
	totalUsers := 100
	var users []*models.User
	headers := []interface{}{"Number", "Name", "Email", "Gender", "Phone", "City", "State", "Country"}

	// ... (user generation logic)
	for i := 0; i < totalUsers; i++ {
		user := models.User{
			Name:   gofakeit.Name(),
			Email:  gofakeit.Email(),
			Gender: models.Gender(gofakeit.Gender()),
			Phone:  gofakeit.Phone(),
			Address: models.Address{
				City:    gofakeit.City(),
				State:   gofakeit.State(),
				Country: gofakeit.Country(),
			},
		}
		users = append(users, &user)
	}

	// Create a map to store writers for each sheet
	sheetWriters := make(map[string]*excelize.StreamWriter)

	// Create new Excel file
	xlsx := excelize.NewFile()

	for i, u := range users {
		// Determine sheet name based on index
		limit := 20
		sheetIndex := (i-1)/limit + 1
		sheetName := fmt.Sprintf("Sheet%d", sheetIndex)

		// Jika sheet belum ada, buat sheet baru
		if !sheetExists(xlsx, sheetName) {
			xlsx.NewSheet(sheetName)
		}

		// Check if sheet exists and create writer if needed
		writer, ok := sheetWriters[sheetName]
		if !ok {
			var err error
			writer, err = xlsx.NewStreamWriter(sheetName)
			if err != nil {
				panic("Error creating writer!")
			}
			sheetWriters[sheetName] = writer

			// Set headers only once for the sheet
			header, err := setHeader(writer, headers)
			if err != nil {
				panic(err)
			}

			// set autofilter
			setAutoFilter(xlsx, sheetName, header)
		}

		// Write data to the sheet using existing writer
		value := []interface{}{i, u.Name, u.Email, u.Gender, u.Phone, u.Address.City, u.Address.State, u.Address.Country}
		streamValues(writer, i, limit, value)

	}

	// Flush all writers and save the file
	for _, writer := range sheetWriters {
		if err := writer.Flush(); err != nil {
			panic("Error flushing writer!")
		}
	}

	if err := xlsx.SaveAs("./output.xlsx"); err != nil {
		fmt.Println("Error:", err)
		return
	}

	fmt.Println("Excel file created successfully.")
}

func TestPrint(t *testing.T) {
	var header []string

	df := &File{
		Headers: []interface{}{"Number", "Name", "Email", "Gender", "Phone", "City", "State", "Country"},
	}

	for _, h := range df.Headers {
		head := fmt.Sprintf("%v", h)
		header = append(header, head)
	}

	fmt.Println(header)
}
