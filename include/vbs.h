#pragma once

//#include <stdint.h>

#include "rust/cxx.h"
#include <Windows.h>

enum class ScriptLang : uint8_t;

int init();
int parse_wrapper(rust::Str code, char* output, ScriptLang lang);
rust::String error_to_string(int hr);
int set_variable(rust::Str name, char* val, ScriptLang lang);