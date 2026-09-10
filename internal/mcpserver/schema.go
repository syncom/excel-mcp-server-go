package mcpserver

import "encoding/json"

// The MCP SDK infers an input schema from the handler's struct, but
// jsonschema-go offers no way to declare a `default`, and SPEC 5.6 compares
// defaults. Importing jsonschema-go to build schemas by hand is not an option
// either — it is an indirect dependency, and making it a direct one would break
// the two-require budget of SPEC 3.
//
// So schemas are declared here as plain structs that marshal to JSON Schema.
// mcp.Tool.InputSchema accepts any value that marshals to a valid schema and
// remarshals it internally, which keeps the dependency budget intact and gives
// exact control over the fields SPEC 5.6 checks.

type schema struct {
	Type       string           `json:"type"`
	Properties map[string]*prop `json:"properties"`
	Required   []string         `json:"required,omitempty"`
	Title      string           `json:"title,omitempty"`
}

type prop struct {
	Type  string  `json:"type,omitempty"`
	AnyOf []*prop `json:"anyOf,omitempty"`
	Items *prop   `json:"items,omitempty"`
	Title string  `json:"title,omitempty"`
	// AdditionalProperties is a pointer-free bool because pydantic emits it as
	// `true` for Dict[str, Any]; omitted everywhere else.
	AdditionalProperties bool `json:"additionalProperties,omitempty"`
	// Default is raw JSON so that `false` and `null` survive; a plain `any`
	// with omitempty would drop exactly the defaults SPEC 5.6 cares about
	// (preview_only's false, end_cell's null).
	Default json.RawMessage `json:"default,omitempty"`
}

// def encodes a default value as raw JSON.
func def(v any) json.RawMessage {
	b, err := json.Marshal(v)
	if err != nil {
		panic("mcpserver: undefaultable value: " + err.Error())
	}
	return b
}

// str, boolean, nullableStr and strList build the property shapes the Python
// signatures produce through pydantic.
func str(title string) *prop     { return &prop{Type: "string", Title: title} }
func boolean(title string) *prop { return &prop{Type: "boolean", Title: title} }

// nullableStr mirrors pydantic's rendering of Optional[str]: an anyOf of the
// type and null, with a null default.
func nullableStr(title string) *prop {
	return &prop{
		AnyOf:   []*prop{{Type: "string"}, {Type: "null"}},
		Title:   title,
		Default: def(nil),
	}
}

// nullableOf mirrors pydantic's Optional[T]: anyOf [T, null] with a null
// default.
func nullableOf(t *prop, title string) *prop {
	return &prop{AnyOf: []*prop{t, {Type: "null"}}, Title: title, Default: def(nil)}
}

// object with additionalProperties, as pydantic renders Dict[str, Any].
func objectProp() *prop { return &prop{Type: "object", AdditionalProperties: true} }

func withDefault(p *prop, v any) *prop {
	p.Default = def(v)
	return p
}

// object assembles a tool input schema. Required order follows the Python
// signature, which is also the order pydantic emits.
func object(title string, required []string, props map[string]*prop) *schema {
	return &schema{Type: "object", Properties: props, Required: required, Title: title}
}
