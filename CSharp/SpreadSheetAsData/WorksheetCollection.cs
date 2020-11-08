using System;
using System.Collections;
using System.Collections.Generic;
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
            items.Where(_ => _.Name == sheetName).SingleOrDefault();

        public Worksheet this[int index] => items.ElementAt(index);

        public int Count => items.Count();

        public IEnumerable<string> Keys => items.Select(_ => _.Name);

        public IEnumerable<Worksheet> Values => items;

        public bool ContainsKey(string key) => Keys.Contains(key);

        public bool TryGetValue(string key, out Worksheet value)
        {
            value = items.Where(_ => _.Name == key).SingleOrDefault();
            return value != null;
        }

        public IEnumerator<Worksheet> GetEnumerator() => items.GetEnumerator();

        IEnumerator<KeyValuePair<string, Worksheet>> IEnumerable<KeyValuePair<string, Worksheet>>.GetEnumerator() =>
            items.ToDictionary(_ => _.Name, _ => _).GetEnumerator();
        
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
