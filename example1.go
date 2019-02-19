// +build ignore

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
		local fsObj = create_object("Scripting.FileSystemObject")
		local folder= fsObj:GetFolder("C:\\")
		local files = folder:_get("Files")
		print("count=",files:_get("Count"))
		for f in files:_iter() do
			print(f:_get("Name"))
			f:_release()
		end
		folder:_release()
		files:_release()
		fsObj:_release()
	`)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
	}
}
