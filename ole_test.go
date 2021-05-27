package ole_test

import (
	"strings"
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

	err = L.DoString(`
		local fsObj = create_object("Scripting.FileSystemObject")
		assert(fsObj._release())`)
	if err == nil {
		t.Fatalf("_release() without receiver has to fail.")
	}
	if errStr := err.Error(); !strings.Contains(errStr, "no receiver") {
		t.Fatalf("OBJECT:_release(): %s", errStr)
	}
	// println(err.Error())
}
