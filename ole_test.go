package ole_test

import (
	"testing"

	"github.com/yuin/gopher-lua"
	"github.com/zetamatta/glua-ole"
)

func newL() *lua.LState {
	L := lua.NewState()
	L.SetGlobal("create_object", L.NewFunction(ole.CreateObject))
	L.SetGlobal("to_ole_integer", L.NewFunction(ole.ToOleInteger))
	return L
}

func TestGc(t *testing.T) {
	L := newL()
	defer L.Close()

	err := L.DoString(`
		local fsObj = create_object("Scripting.FileSystemObject")
		fsObj:_release()`)
	if err != nil {
		t.Fatalf("fsObj:_release() failed: %s", err)
	}
}
