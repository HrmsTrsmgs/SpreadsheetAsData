using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace Marimo.SpreadSheetAsData
{
    public class WorksheetCollection : IReadOnlyList<Worksheet>, IReadOnlyDictionary<string, Worksheet>
    {
        internal WorksheetCollection(IEnumerable<Worksheet> collection)
        {
            items = collection.ToArray();
        }

        IEnumerable<Worksheet> items { get; }

        public Worksheet this[string sheetName] =>
            items.Where(_ => _.Name == sheetName).SingleOrDefault()
                ?? throw new KeyNotFoundException();

        public Worksheet this[int index] => items.ElementAt(index);

        public int Count => items.Count();

        public IEnumerable<string> Keys => items.Select(_ => _.Name);

        public IEnumerable<Worksheet> Values => items;

        public bool ContainsKey(string key) => Keys.Contains(key);

        public bool TryGetValue(string key, [MaybeNullWhen(false)] out Worksheet value)
        {
            var sheet = items.Where(_ => _.Name == key).SingleOrDefault();
            value = sheet;
            return sheet != null;
        }

        public IEnumerator<Worksheet> GetEnumerator() => items.GetEnumerator();

        IEnumerator<KeyValuePair<string, Worksheet>> IEnumerable<KeyValuePair<string, Worksheet>>.GetEnumerator() =>
            items.ToDictionary(_ => _.Name, _ => _).GetEnumerator();
        
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
