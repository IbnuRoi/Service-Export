package db

import (
	"database/sql"
	"fmt"
	"sam/RnD/codebase/service-export/config"
)

var db map[string]*sql.DB

func Connect(dbName string) {
	if val, ok := config.GetConfig().DatabaseConfig[dbName]; !ok {
		panic(fmt.Errorf("error initialize db %v", dbName))
	}

}

func GetDb(dbName string) *sql.DB {
	return db[dbName]
}
