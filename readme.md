glua-ole 
========

The bridge library between [GopherLua](https://github.com/yuin/gopher-lua)
and [go-ole](https://github.com/go-ole/go-ole).

Using
------

```go
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
		local files = fsObj:GetFolder("C:\\"):_get("Files")
		print(files:_get("Count"))
	`)
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
	}
}
```

- `local OBJ=create_object()` creates OLE-Object
- `OBJ:method(...)` calls method
- `OBJ:_get("PROPERTY")` returns the value of the property.
- `OBJ:_set("PROPERTY",value)` sets the value to the property.
- `local N=to_ole_integer(10)` creates the integer value for OLE.
