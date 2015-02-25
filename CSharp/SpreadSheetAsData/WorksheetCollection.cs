using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public class WorksheetCollection : IReadOnlyList<Worksheet>, IReadOnlyDictionary<string, Worksheet>
    {
        internal WorksheetCollection(IEnumerable<Worksheet> collection)
        {
            items = collection.ToList();
        }

        List<Worksheet> items;

        public Worksheet this[string sheetName]
        {
            get
            {
                return items.Where(_ => _.Name == sheetName).SingleOrDefault();
            }
        }

        public Worksheet this[int index]
        {
            get
            {
                return items[index];
            }
        }

        public int Count
        {
            get
            {
                return items.Count;
            }
        }

        public IEnumerable<string> Keys
        {
            get
            {
                return items.Select(_ => _.Name);
            }
        }

        public IEnumerable<Worksheet> Values
        {
            get
            {
                return items;
            }
        }

        public bool ContainsKey(string key)
        {
            return Keys.Contains(key);
        }

        public bool TryGetValue(string key, out Worksheet value)
        {
            value = items.Where(_ => _.Name == key).SingleOrDefault();
            return value != null;
        }

        public IEnumerator<Worksheet> GetEnumerator()
        {
            return items.GetEnumerator();
        }

        IEnumerator<KeyValuePair<string, Worksheet>> IEnumerable<KeyValuePair<string, Worksheet>>.GetEnumerator()
        {
            return items.ToDictionary(_ => _.Name, _ => _).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return items.GetEnumerator();
        }
    }
}
