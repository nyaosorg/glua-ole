// +build ignore

// Use `go run example2.go`
// This program tests whether COM leak exists or not by EXCEL

package main

import (
	"fmt"
	"os"

	"github.com/yuin/gopher-lua"
	"github.com/zetamatta/glua-ole"
)

func main() {
	L := lua.NewState()
	defer L.Close()

	L.SetGlobal("create_object", L.NewFunction(ole.CreateObject))
	L.SetGlobal("to_ole_integer", L.NewFunction(ole.ToOleInteger))

	err := L.DoString(`
		local excel = create_object("Excel.Application")
		excel:_set("Visible",true)
		excel:Quit()
		excel:_release()
	`)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
	}
}
