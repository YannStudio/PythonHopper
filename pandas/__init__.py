class DataFrame:
    def __init__(self, data=None, columns=None):
        self._rows = []
        self.columns = []
        if data is None:
            data = []
        if isinstance(data, dict):
            self.columns = list(data.keys())
            rows = zip(*data.values())
            for row in rows:
                self._rows.append(dict(zip(self.columns, row)))
        elif data and isinstance(data[0], dict):
            cols = set()
            for row in data:
                cols.update(row.keys())
            self.columns = list(cols)
            for row in data:
                self._rows.append(dict(row))
        elif data:
            self.columns = list(columns or [])
            for row in data:
                self._rows.append(dict(zip(self.columns, row)))
        else:
            self.columns = list(columns or [])

    def iterrows(self):
        for idx, row in enumerate(self._rows):
            yield idx, row

    def to_dict(self, orient="records"):
        if orient == "records":
            return [dict(r) for r in self._rows]
        raise NotImplementedError

    def __getitem__(self, key):
        return [r.get(key) for r in self._rows]

    def __setitem__(self, key, value):
        if not isinstance(value, list):
            value = [value] * len(self._rows)
        if len(self._rows) < len(value):
            for _ in range(len(value) - len(self._rows)):
                self._rows.append({})
        for row, val in zip(self._rows, value):
            row[key] = val
        if key not in self.columns:
            self.columns.append(key)

    def __len__(self):
        return len(self._rows)

# minimal namespace
def read_csv(*args, **kwargs):
    raise NotImplementedError("read_csv is not implemented in stub pandas")

def read_excel(*args, **kwargs):
    raise NotImplementedError("read_excel is not implemented in stub pandas")
