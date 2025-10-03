//go:build windows
// +build windows

package wmi

import (
	"errors"
	"unsafe"

	ole "github.com/go-ole/go-ole"
	"golang.org/x/sys/windows"
)

var (
	success = errors.New("The operation completed successfully.")
)

var (
	modoleaut32               = windows.NewLazySystemDLL("oleaut32.dll")
	procSafeArrayCreateVector = modoleaut32.NewProc("SafeArrayCreateVector")
	procSafeArrayPutElement   = modoleaut32.NewProc("SafeArrayPutElement")
)

// safeArrayCreateVector creates SafeArray.
//
// AKA: SafeArrayCreateVector in Windows API.
func safeArrayCreateVector(variantType ole.VT, lowerBound int32, length uint32) (safearray *ole.SafeArray, err error) {
	sa, _, err := procSafeArrayCreateVector.Call(
		uintptr(variantType),
		uintptr(lowerBound),
		uintptr(length),
	)
	if !(errors.Is(err, success) || err.Error() == success.Error()) { // errors.Is not working as expected
		return nil, err
	}
	saPtr := (**uintptr)(unsafe.Pointer(&sa)) // create Pointer of SafeArray-Pointer to solve possible misuse of unsafe.Pointer warning
	return (*ole.SafeArray)(unsafe.Pointer(*saPtr)), nil
}

// safeArrayPutElement stores the data element at the specified location in the
// array.
//
// AKA: SafeArrayPutElement in Windows API.
func safeArrayPutElement(safearray *ole.SafeArray, index int64, element uintptr) (err error) {
	err = convertHresultToError(
		procSafeArrayPutElement.Call(
			uintptr(unsafe.Pointer(safearray)),
			uintptr(unsafe.Pointer(&index)),
			uintptr(unsafe.Pointer(element))))
	return
}

// convertHresultToError converts syscall to error, if call is unsuccessful.
func convertHresultToError(hr uintptr, r2 uintptr, ignore error) (err error) {
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}
