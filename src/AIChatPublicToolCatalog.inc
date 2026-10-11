// 公共工具的固定定义；JSON Unicode 转义保证 UTF-8 语义与执行字符集无关。
R"json(
[
  {
    "name": "refresh_workspace_mirror",
    "description": "Refresh the current e-packager workspace mirror from live IDE memory. External MCP clients MUST call this successfully before their first source read/edit tool in each MCP session. Built-in AI chat automatically performs a full refresh before each request, while this tool remains available for explicit retries or later refreshes. mode=auto keeps the default strategy, main_only refreshes only main source files when possible, and full rebuilds the complete mirror.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "mode": {
          "type": "string",
          "enum": [
            "auto",
            "main_only",
            "full"
          ],
          "description": "Defaults to auto."
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "e_packager",
    "description": "Only call this tool when the user explicitly asks to parse, inspect, reference, or reproduce an external .e/.ec file. It unpacks that file with e-packager into unimported_code/{file name}/ inside the current workspace mirror. The unpacked files are read-only reference code: use list_files or search_code to locate them, then read_file/read_files to read their contents. Relative file_path values resolve from the current project's directory; existing output for the same file name is replaced only after unpack succeeds.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string",
          "minLength": 1,
          "maxLength": 32768,
          "description": "Absolute path to an existing .e/.ec file, or a path relative to the current project directory."
        }
      },
      "required": [
        "file_path"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "update_plan",
    "description": "Update the visible task plan for routine work whenever it has two or more meaningful steps or progress tracking would help. Prefer it for everyday implementation, diagnosis, and verification; skip only trivial one-step actions. It does not modify source code and does not replace the <proposed_plan> approval flow in plan mode.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "explanation": {
          "type": "string",
          "description": "Optional short note about why the plan changed."
        },
        "plan": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "step": {
                "type": "string"
              },
              "status": {
                "type": "string",
                "enum": [
                  "pending",
                  "in_progress",
                  "completed"
                ]
              }
            },
            "required": [
              "step",
              "status"
            ],
            "additionalProperties": false
          },
          "description": "Task steps. At most one item may be in_progress."
        }
      },
      "required": [
        "plan"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "read_skill_resource",
    "description": "Read SKILL.md or a text resource inside one enabled AutoLinker skill. Call this before using a skill that clearly matches the task but was not explicitly injected with $skill-name. Relative paths are restricted to the selected skill directory; continue large resources with next_byte_offset.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "skill_name": {
          "type": "string",
          "minLength": 1,
          "maxLength": 128
        },
        "relative_path": {
          "type": "string",
          "maxLength": 2048,
          "description": "Defaults to SKILL.md."
        },
        "byte_offset": {
          "type": "integer",
          "minimum": 0
        },
        "max_bytes": {
          "type": "integer",
          "minimum": 4096,
          "maximum": 262144
        }
      },
      "required": [
        "skill_name"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "list_files",
    "description": "List files in the current e-packager workspace mirror. Paths are relative to the mirror root. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. By default focuses on src, ecom, elib, header and unimported_code text areas. Continue paginated results with next_offset.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "glob": {
          "type": "string",
          "description": "Optional glob such as src/**/*.txt or ecom/**/*.txt."
        },
        "path": {
          "type": "string",
          "description": "Optional relative path prefix."
        },
        "mirror_generation": {
          "type": "integer",
          "minimum": 1,
          "description": "Pagination continuation only. MUST omit when offset is 0 or no previous page generation is available; never guess a value such as 1. For offset > 0, pass the exact generation returned by the previous page."
        },
        "offset": {
          "type": "integer",
          "minimum": 0,
          "description": "Zero-based result offset for pagination. Use next_offset from the prior result."
        },
        "limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 5000
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "search_code",
    "description": "Batch-stream search complete text files inside the current e-packager workspace mirror. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. Literal substring search is the default; set regex=true only for simple bounded regular expressions. Use patterns to match up to 16 related names in one file pass.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "pattern": {
          "type": "string",
          "maxLength": 1024,
          "description": "Single literal substring by default; set regex=true for regular-expression search."
        },
        "patterns": {
          "type": "array",
          "maxItems": 16,
          "items": {
            "type": "string",
            "maxLength": 1024
          },
          "description": "Optional batch of patterns using the same regex/case/glob/context options."
        },
        "glob": {
          "type": "string",
          "description": "Optional file glob filter such as src/**/*.txt."
        },
        "output_mode": {
          "type": "string",
          "enum": [
            "files_with_matches",
            "content",
            "count"
          ]
        },
        "regex": {
          "type": "boolean",
          "description": "Defaults to false."
        },
        "case_insensitive": {
          "type": "boolean"
        },
        "context": {
          "type": "integer",
          "minimum": 0,
          "maximum": 20
        },
        "mirror_generation": {
          "type": "integer",
          "minimum": 1,
          "description": "Pagination continuation only. MUST omit when offset is 0 or no previous search page generation is available; never guess a value such as 1. For offset > 0, pass the exact generation returned by the previous page."
        },
        "offset": {
          "type": "integer",
          "minimum": 0,
          "description": "Zero-based result offset. Continue with next_offset returned by the previous page."
        },
        "head_limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 2000
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "read_file",
    "description": "Read one text file from the current e-packager workspace mirror. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. Returns numbered text, line pagination, mirror_generation and a 1 MiB source byte window. Continue large files with next_source_byte_offset as byte_offset.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "mirror_generation": {
          "type": "integer",
          "minimum": 1,
          "description": "Pagination continuation only. MUST omit when offset and byte_offset are 0 or no previous read generation is available; never guess. For continuation, pass the exact generation returned by the previous page."
        },
        "byte_offset": {
          "type": "integer",
          "minimum": 0,
          "description": "Source byte cursor for files larger than the 1 MiB read window. Continue with next_source_byte_offset and reset line offset to 0."
        },
        "offset": {
          "type": "integer",
          "minimum": 0,
          "description": "Zero-based line offset. Positive initial offsets are allowed; include mirror_generation when continuing a prior page for cross-refresh consistency."
        },
        "limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 20000
        }
      },
      "required": [
        "file_path"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "read_files",
    "description": "Batch read one or more text files or byte windows from the current e-packager workspace mirror. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. Prefer this over repeated read_file calls. Partial failures are explicit and large files continue with per-file next_source_byte_offset.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_paths": {
          "type": "array",
          "maxItems": 12,
          "items": {
            "type": "string"
          },
          "description": "Simple list of up to 12 mirror-relative file paths."
        },
        "mirror_generation": {
          "type": "integer",
          "minimum": 1,
          "description": "Pagination continuation only. MUST omit for an initial batch read or when no previous read generation is available; never guess. For continuation, pass the exact generation returned by the previous page."
        },
        "files": {
          "type": "array",
          "maxItems": 12,
          "items": {
            "type": "object",
            "properties": {
              "file_path": {
                "type": "string"
              },
              "mirror_generation": {
                "type": "integer",
                "minimum": 1,
                "description": "Compatibility alias for the top-level mirror_generation. If present in multiple entries, all values must match."
              },
              "byte_offset": {
                "type": "integer",
                "minimum": 0
              },
              "offset": {
                "type": "integer",
                "minimum": 0
              },
              "limit": {
                "type": "integer",
                "minimum": 1,
                "maximum": 2000
              }
            },
            "required": [
              "file_path"
            ],
            "additionalProperties": false
          },
          "description": "Optional per-file offset/limit entries, up to 12 per call."
        },
        "byte_offset": {
          "type": "integer",
          "minimum": 0,
          "description": "Default source byte cursor for all entries."
        },
        "offset": {
          "type": "integer",
          "minimum": 0,
          "description": "Default offset for file_paths/files entries."
        },
        "limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 2000,
          "description": "Default per-file line limit."
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "read_code_item",
    "description": "Stream-locate and read one complete E-language top-level code item (especially a .\u5b50\u7a0b\u5e8f), including items beyond the first 1 MiB. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. Use occurrence to disambiguate duplicate names. Optional include_references adds a bounded case-insensitive literal reference search.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "item_name": {
          "type": "string"
        },
        "occurrence": {
          "type": "integer",
          "minimum": 1,
          "description": "Select the Nth matching declaration when duplicate top-level names exist."
        },
        "mirror_generation": {
          "type": "integer",
          "minimum": 1,
          "description": "Pagination continuation only. MUST omit when offset is 0 or no prior generation is available; never guess. For continuation, pass the exact generation returned by the previous result."
        },
        "include_references": {
          "type": "boolean",
          "description": "Defaults to false."
        },
        "reference_glob": {
          "type": "string",
          "description": "Optional reference search glob such as src/**/*.txt."
        },
        "reference_limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 50
        }
      },
      "required": [
        "file_path",
        "item_name"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "read_real_file",
    "description": "Read a bounded numbered view directly from the live IDE page mapped by mirror-relative file_path. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. Returns code_hash plus paginated content; continue with next_offset when has_more=true. Call it immediately before editing and base old_text/full_code/expected_base_hash on this live real_source view.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "offset": {
          "type": "integer",
          "minimum": 0
        },
        "limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 20000
        }
      },
      "required": [
        "file_path"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "edit_file",
    "description": "Edit one current-project source file by mirror-relative file_path. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. The edit is always based on the live IDE page; expected_base_hash optionally rejects a stale caller base. Successful results are verified.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "old_text": {
          "type": "string",
          "description": "Exact match text. ASCII and IDE-style left/right double quotes are treated as equivalent."
        },
        "new_text": {
          "type": "string"
        },
        "expected_base_hash": {
          "type": "string",
          "description": "code_hash from read_real_file; required for external MCP calls and rejects stale live source."
        }
      },
      "required": [
        "file_path",
        "old_text",
        "new_text"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "multi_edit_file",
    "description": "Apply multiple text edits to one current-project source file. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. The edits are always based on the live IDE page; expected_base_hash optionally rejects stale live source.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "edits": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "old_text": {
                "type": "string",
                "description": "Exact match text. ASCII and IDE-style left/right double quotes are treated as equivalent."
              },
              "new_text": {
                "type": "string"
              },
              "replace_all": {
                "type": "boolean"
              }
            },
            "required": [
              "old_text",
              "new_text"
            ],
            "additionalProperties": false
          }
        },
        "fail_on_unmatched": {
          "type": "boolean"
        },
        "atomic": {
          "type": "boolean"
        },
        "expected_base_hash": {
          "type": "string",
          "description": "code_hash from read_real_file; required for external MCP calls and rejects stale live source."
        }
      },
      "required": [
        "file_path",
        "edits"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "write_file",
    "description": "Overwrite one current-project source file with full_code. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. expected_base_hash detects stale source. Successful results include verified=true and do not require confirmation reads.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "full_code": {
          "type": "string"
        },
        "expected_base_hash": {
          "type": "string",
          "description": "code_hash from read_real_file; required for external MCP calls."
        }
      },
      "required": [
        "file_path",
        "full_code"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "diff_file",
    "description": "Preview a structured diff for one current-project source file without writing anything. The diff is based on the live IDE page. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. Accepts new_code/full_code or text edit parameters.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "new_code": {
          "type": "string"
        },
        "full_code": {
          "type": "string"
        },
        "old_text": {
          "type": "string",
          "description": "Exact match text. ASCII and IDE-style left/right double quotes are treated as equivalent."
        },
        "new_text": {
          "type": "string"
        },
        "edits": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "old_text": {
                "type": "string",
                "description": "Exact match text. ASCII and IDE-style left/right double quotes are treated as equivalent."
              },
              "new_text": {
                "type": "string"
              },
              "replace_all": {
                "type": "boolean"
              }
            },
            "required": [
              "old_text",
              "new_text"
            ],
            "additionalProperties": false
          }
        },
        "fail_on_unmatched": {
          "type": "boolean"
        },
        "expected_base_hash": {
          "type": "string",
          "description": "code_hash from read_real_file; required for external MCP calls and rejects stale live source."
        }
      },
      "required": [
        "file_path"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "restore_file_snapshot",
    "description": "Restore one current-project source file from the latest real-page snapshot or a specified snapshot_id. Requires a prepared mirror; external MCP calls refresh_workspace_mirror, while built-in chat prepares it automatically. The current live page is read first; expected_current_hash optionally prevents overwriting newer changes.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string"
        },
        "snapshot_id": {
          "type": "string"
        },
        "restore_latest": {
          "type": "boolean"
        },
        "expected_current_hash": {
          "type": "string",
          "description": "Current live code hash; required for external MCP restore calls."
        }
      },
      "required": [
        "file_path"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "add_new_file",
    "description": "Create a new assembly or class in the current E-language project, optionally write its complete source, verify the final IDE object name, and rebuild the full workspace mirror. full_code may include a .\u7a0b\u5e8f\u96c6 header; its name is normalized to the requested name.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_type": {
          "type": "string",
          "enum": [
            "assembly",
            "class"
          ],
          "description": "Create a normal assembly or a class module."
        },
        "name": {
          "type": "string",
          "minLength": 1,
          "maxLength": 255,
          "description": "Final assembly or class name."
        },
        "full_code": {
          "type": "string",
          "description": "Optional complete E-language page source. If omitted, the IDE-generated default body is retained and only the header name changes."
        }
      },
      "required": [
        "file_type",
        "name"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "delete_file",
    "description": "Delete one local assembly or class and all its subroutines from the live E-language project using the native IDE command, verify its removal, and rebuild the workspace mirror. Read read_real_file first and pass code_hash as expected_base_hash. Does not rewrite references in other pages. Window-bound assemblies, form XML, fixed tables and imported module files are rejected. Does not save the project to disk.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "file_path": {
          "type": "string",
          "minLength": 1,
          "description": "Exact workspace mirror source path returned by list_files or add_new_file."
        },
        "expected_base_hash": {
          "type": "string",
          "minLength": 1,
          "description": "Live source code_hash from read_real_file; deletion is rejected if the page has changed."
        }
      },
      "required": [
        "file_path",
        "expected_base_hash"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "get_current_page_info",
    "description": "Get current IDE page name, page type and the trace/source used to resolve that page name.",
    "inputSchema": {
      "type": "object",
      "properties": {},
      "additionalProperties": false
    }
  },
  {
    "name": "get_current_eide_info",
    "description": "Get current E-language IDE instance information, including current source file path, current page info, MCP port/endpoint, process id, executable path and supported compile modes.",
    "inputSchema": {
      "type": "object",
      "properties": {},
      "additionalProperties": false
    }
  },
  {
    "name": "refresh_dependency_catalog",
    "description": "Explicit-user-request only. Refresh AutoLinker dependency catalog cache for available .ec modules under ecom and .fne support libraries under lib.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "force": {
          "type": "boolean",
          "description": "Rebuild cache directories. Defaults to false."
        },
        "wait": {
          "type": "boolean",
          "description": "Wait until refresh finishes. Defaults to false so the IDE main thread is not blocked."
        },
        "timeout_ms": {
          "type": "integer",
          "minimum": 0,
          "maximum": 600000,
          "description": "Defaults to 30000 when wait=true; 0 also uses the bounded default."
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "search_available_modules",
    "description": "Explicit-user-request only. Search currently available .ec modules from the E-language ecom directory by module name, path, or cached main source.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "query": {
          "type": "string",
          "description": "Module name, file name, path or source keyword. Empty query lists top cached modules."
        },
        "limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 200
        },
        "include_snippets": {
          "type": "boolean"
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "search_available_support_libraries",
    "description": "Explicit-user-request only. Search currently available .fne support libraries from the E-language lib directory by library name, file name, path, or decoded GetNewInf information.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "query": {
          "type": "string",
          "description": "Support library name, file name, path, command name or type keyword. Empty query lists top cached libraries."
        },
        "limit": {
          "type": "integer",
          "minimum": 1,
          "maximum": 200
        },
        "include_snippets": {
          "type": "boolean"
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "list_imported_modules",
    "description": "Explicit-user-request only. List .ec modules currently imported by the active E-language project.",
    "inputSchema": {
      "type": "object",
      "properties": {},
      "additionalProperties": false
    }
  },
  {
    "name": "add_module_to_project",
    "description": "Explicit-user-request only. Import one .ec module into the current project by module_path or module_name. Real import requires allow_blocking_import=true.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "module_path": {
          "type": "string",
          "description": "Absolute .ec file path."
        },
        "module_name": {
          "type": "string",
          "description": "Module name or file name to resolve from the E-language module directory."
        },
        "prefer_new_method": {
          "type": "boolean",
          "description": "Optional diagnostic switch. false uses the legacy IDE import path; true tries the newer AddECOM2 path first."
        },
        "allow_blocking_import": {
          "type": "boolean",
          "description": "Required true to execute the real IDE module import call. Default false only resolves the module and avoids blocking MCP if the IDE API hangs."
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "remove_module_from_project",
    "description": "Explicit-user-request only. Remove one imported .ec module from the current project by module_index, module_path, or module_name.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "module_index": {
          "type": "integer",
          "minimum": 0
        },
        "module_path": {
          "type": "string"
        },
        "module_name": {
          "type": "string"
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "add_support_library_to_project",
    "description": "Explicit-user-request only. Add one .fne support library to the current project by library_path or library_name, then verify it is loaded.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "library_path": {
          "type": "string",
          "description": "Absolute .fne file path, or path relative to the E-language lib directory."
        },
        "library_name": {
          "type": "string",
          "description": "Support library name or file name to resolve from the E-language lib directory."
        }
      },
      "additionalProperties": false
    }
  },
  {
    "name": "compile_with_output_path",
    "description": "Compile the current project to output_path. compiler defaults to eide: omit blackmoon_mode and optionally set static_compile (default true for EXE/DLL). For eide module projects, auto builds a win_console_exe test; target=ecom publishes .ec. BlackMoon requires compiler=blackmoon, an EXE/DLL project, and an available profile; omit static_compile entirely, set blackmoon_mode=asm/cxx/vcxx (default vcxx), and use target=auto or the actual project type. A stray blackmoon_mode never switches the compiler: conflicting parameters return error and hint explaining how to correct the call. Success requires an updated output artifact and successful configuration restoration.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "target": {
          "type": "string",
          "enum": [
            "auto",
            "win_exe",
            "win_console_exe",
            "win_dll",
            "ecom"
          ],
          "description": "Defaults to auto. For module-project compile verification, use win_console_exe; ecom publishes a .ec module."
        },
        "output_path": {
          "type": "string"
        },
        "static_compile": {
          "type": "boolean",
          "description": "Only for compiler=eide: defaults to true for EXE/DLL and false for ecom. For compiler=blackmoon omit this field entirely, including false."
        },
        "compiler": {
          "type": "string",
          "enum": [
            "eide",
            "blackmoon"
          ]
        },
        "blackmoon_mode": {
          "type": "string",
          "enum": [
            "asm",
            "cxx",
            "vcxx"
          ],
          "description": "Only for explicit compiler=blackmoon; omit for ordinary eide compilation. Default vcxx. asm=\u6c47\u7f16\u6a21\u5f0f, cxx=C/C++\u6a21\u5f0f, vcxx=VC++\u6a21\u5f0f. Use get_current_eide_info to inspect installed modes."
        }
      },
      "required": [
        "output_path"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "exec_command",
    "description": "Runs a command in a non-interactive local shell and returns output or a session ID for ongoing polling. The command starts only after explicit user confirmation. tty=true is not supported by AutoLinker; use write_stdin with an empty chars value to poll a running session.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "cmd": {
          "type": "string",
          "description": "Shell command to execute."
        },
        "workdir": {
          "type": "string",
          "description": "Working directory. Defaults to the current e-language project directory."
        },
        "tty": {
          "type": "boolean",
          "description": "Only false is supported; AutoLinker does not provide ConPTY."
        },
        "yield_time_ms": {
          "type": "integer",
          "description": "Wait before yielding output. Defaults to 10000 ms; effective range is 250-30000 ms, with a 2000 ms Windows initial floor."
        },
        "max_output_tokens": {
          "type": "integer",
          "description": "Output token budget. Defaults to 10000; larger requests are capped by policy."
        },
        "shell": {
          "type": "string",
          "description": "Shell executable. Defaults to powershell.exe."
        },
        "login": {
          "type": "boolean",
          "description": "Use login/profile shell semantics. Defaults to true."
        }
      },
      "required": [
        "cmd"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "write_stdin",
    "description": "Writes characters to an existing non-TTY exec_command session and returns recent output. Empty chars polls the session; only Ctrl-C is accepted as non-empty input because AutoLinker uses pipes rather than ConPTY.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "session_id": {
          "type": "integer",
          "description": "Identifier of the running exec_command session."
        },
        "chars": {
          "type": "string",
          "description": "Bytes to write. Defaults to empty, which polls without writing."
        },
        "yield_time_ms": {
          "type": "integer",
          "description": "Non-empty writes default to 250 ms and cap at 30000 ms; empty polls wait 5000-300000 ms."
        },
        "max_output_tokens": {
          "type": "integer",
          "description": "Output token budget. Defaults to 10000; larger requests are capped by policy."
        }
      },
      "required": [
        "session_id"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "search_web_tavily",
    "description": "Search the public web via Tavily and return normalized result snippets.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "query": {
          "type": "string"
        },
        "max_results": {
          "type": "integer",
          "minimum": 1,
          "maximum": 10
        },
        "topic": {
          "type": "string"
        }
      },
      "required": [
        "query"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "fetch_url",
    "description": "Fetch one URL via HTTP GET and return normalized text response plus basic HTTP metadata.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "url": {
          "type": "string"
        },
        "timeout_seconds": {
          "type": "integer",
          "minimum": 1,
          "maximum": 300
        },
        "max_bytes": {
          "type": "integer",
          "minimum": 4096,
          "maximum": 2097152
        }
      },
      "required": [
        "url"
      ],
      "additionalProperties": false
    }
  },
  {
    "name": "extract_web_document",
    "description": "Fetch one web page or text document and extract readable plain-text content, title and a small set of links.",
    "inputSchema": {
      "type": "object",
      "properties": {
        "url": {
          "type": "string"
        },
        "timeout_seconds": {
          "type": "integer",
          "minimum": 1,
          "maximum": 300
        },
        "max_bytes": {
          "type": "integer",
          "minimum": 4096,
          "maximum": 2097152
        }
      },
      "required": [
        "url"
      ],
      "additionalProperties": false
    }
  }
]
)json"
