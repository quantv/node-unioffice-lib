#!/bin/bash
set -e
go build -buildmode=c-shared -ldflags="-s -w" -o "dist/libspreadsheet.so" src/lib.go src/funcs.go

rm dist/libspreadsheet.zip
zip -j dist/libspreadsheet.zip dist/*
#go build -buildmode=c-shared -o "include/libspreadsheet.so" src-go/lib.go src-go/funcs.go
#ldflags="-s -w"
