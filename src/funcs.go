package main

/*
#include <errno.h>

void set_err(int err) {
	errno = err;
}
*/
import "C"

func main() {
}

func setError(err int) {
	C.set_err(C.int(err))
}
