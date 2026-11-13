/// Small helpers for building deterministic JSON documents.

/// A JSON string value.
pub fn jstr(value : String) -> @json.Json {
  @json.Json::string(value)
}

/// A JSON number from an integer.
pub fn jint(value : Int) -> @json.Json {
  @json.Json::number(value.to_double())
}

/// A JSON number from a double, emitting `null` for non-finite values.
pub fn jnum(value : Double) -> @json.Json {
  if value.is_nan() || value.is_inf() {
    @json.Json::null()
  } else {
    @json.Json::number(value)
  }
}

/// A JSON boolean.
pub fn jbool(value : Bool) -> @json.Json {
  @json.Json::boolean(value)
}

/// JSON `null`.
pub fn jnull() -> @json.Json {
  @json.Json::null()
}

/// An optional string as JSON (missing values become `null`).
pub fn jopt_str(value : String?) -> @json.Json {
  match value {
    Some(v) => @json.Json::string(v)
    None => @json.Json::null()
  }
}

/// An optional integer as JSON.
pub fn jopt_int(value : Int?) -> @json.Json {
  match value {
    Some(v) => @json.Json::number(v.to_double())
    None => @json.Json::null()
  }
}

/// An optional double as JSON.
pub fn jopt_num(value : Double?) -> @json.Json {
  match value {
    Some(v) => jnum(v)
    None => @json.Json::null()
  }
}

/// A JSON array.
pub fn jarr(items : Array[@json.Json]) -> @json.Json {
  @json.Json::array(items)
}

/// A JSON object with keys sorted so that output is stable across runs.
pub fn jobj(pairs : Array[(String, @json.Json)]) -> @json.Json {
  let sorted = pairs.copy()
  sorted.sort_by(fn(a, b) { a.0.compare(b.0) })
  let map : Map[String, @json.Json] = Map([])
  for pair in sorted {
    map.set(pair.0, pair.1)
  }
  @json.Json::object(map)
}

/// A JSON array of strings.
pub fn jstr_array(items : Array[String]) -> @json.Json {
  let out : Array[@json.Json] = []
  for item in items {
    out.push(@json.Json::string(item))
  }
  @json.Json::array(out)
}

/// Increments a string counter.
pub fn bump(counts : Map[String, Int], key : String) -> Unit {
  let current = match counts.get(key) {
    Some(v) => v
    None => 0
  }
  counts.set(key, current + 1)
}

/// The `limit` most frequent keys, ties broken alphabetically.
pub fn top_entries(counts : Map[String, Int], limit : Int) -> Array[(String, Int)] {
  let items : Array[(String, Int)] = []
  for key, value in counts {
    items.push((key, value))
  }
  items.sort_by(fn(a, b) {
    if a.1 != b.1 {
      return b.1 - a.1
    }
    a.0.compare(b.0)
  })
  let out : Array[(String, Int)] = []
  let mut i = 0
  while i < items.length() && i < limit {
    out.push(items[i])
    i = i + 1
  }
  out
}

/// Total number of observations in a counter.
pub fn total(counts : Map[String, Int]) -> Int {
  let mut sum = 0
  for _, value in counts {
    sum = sum + value
  }
  sum
}
