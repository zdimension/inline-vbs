#pragma once

#include "rust/cxx.h"
#include <Windows.h>

int init();
int parse_wrapper(rust::Str code, char* output);
rust::String error_to_string(int hr);