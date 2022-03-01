#pragma once

#include "rust/cxx.h"
#include <Windows.h>

int init();
int parse(rust::Str code);
int close();
rust::String error_to_string(int hr);