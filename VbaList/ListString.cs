using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VbaList {

    [Description("The type of numbering to use for " + nameof(ListString.GetAsNumberedList) + "() ... if more than 26 items " +
            "then will always use " + nameof(ListType) + "_" + nameof(DIGITS) + ".")]
    [Guid("6EA115F7-0CB3-443F-AE46-1BFA3AF06C26")]
    public enum ListType {
        [Description("Numbering using digits.")]
        DIGITS,
        [Description("Numbering using lowercase letters.")]
        LOWERCASE,
        [Description("Numbering using uppercase letters.")]
        UPPERCASE
    }


    [Guid("160F40F9-11DC-4351-BEFF-317D6A757995")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // default - allow early and late binding
    public interface IListString {

        [Description("Add a String to the end of the " + nameof(ListString) + ".")]
        [DispId(101)]
        void Add(string item);

        [Description("Get the String at a specific index (zero-based) of the " + nameof(ListString) + ".")]
        [DispId(102)]
        string Get(int index);

        [Description("Get the number of Strings held in the " + nameof(ListString) + ".")]
        [DispId(103)]
        int Count();

        [Description("Sort the Strings in the " + nameof(ListString) + ".")]
        [DispId(104)]
        void Sort(bool ascending = true, bool ignoreCase = false);

        [Description("Reverse the order of Strings in the " + nameof(ListString) + ".")]
        [DispId(105)]
        void Reverse();

        [Description("Make every String in the " + nameof(ListString) + " unique by removing duplicates.")]
        [DispId(106)]
        void MakeUnique(bool ignoreCase = false);

        [Description("Get a " + nameof(ListIterator) + " to iterate over the " + nameof(ListString) + ".")]
        [DispId(107)]
        ListIterator GetIterator();

        [Description("Get all Strings in the " + nameof(ListString) + " as a single line of text.")]
        [DispId(108)]
        string GetAsText(string sep = ", ", string lastSep = " and ", string ifNone = "none", string surround = "");

        [Description("Get all Strings in the " + nameof(ListString) + " as a bulleted list.")]
        [DispId(109)]
        string GetAsBullettedList(string bullet = "\u2022 ", string sep = "", string lastSep = "", string ifNone = "",
                string surround = "");

        [Description("Add an Array of Strings to the end of the " + nameof(ListString) + ". The Array must have a 0 lower " +
                "bound and can be either static or dynamic.")]
        [DispId(110)]
        void AddArray(ref string[] strings);

        [Description("Return the " + nameof(ListString) + " as an Array (can either be assigned to a dynamic Array of Strings " +
                "or to a Variant.")]
        [DispId(111)]
        string[] ToArray();

        [Description("Insert a String into any valid index (zero-based) of the " + nameof(ListString) + ".")]
        [DispId(112)]
        void Insert(int index, string item);

        [Description("Find the index of the first matching String in the " + nameof(ListString) + ", search from start to " +
                "end, optionally specifying an initial index to search from and a count of Strings to search. Indexes are zero-based.")]
        [DispId(113)]
        int IndexOf(string item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue);

        [Description("Clear all Strings from the " + nameof(ListString) + ".")]
        [DispId(114)]
        void Clear();

        [Description("Does the " + nameof(ListString) + " contain the given String.")]
        [DispId(115)]
        bool Contains(string text, bool ignoreCase = false);

        [Description("Find the index of the last matching String in the " + nameof(ListString) + ", search from end to start, " +
                "optionally specifying an initial index to search from and a count of Strings to search. Indexes are zero-based.")]
        [DispId(116)]
        int LastIndexOf(string item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue);

        [Description("Get all Strings in the " + nameof(ListString) + " as a numbered list.")]
        [DispId(117)]
        string GetAsNumberedList(ListType listType, string append = ". ", string sep = "", string lastSep = "", string ifNone = "",
                string surround = "");

        [Description("Supports use of \"For Each ... In\" (do not use this method directly).")]
        [DispId(-4)]
        IEnumerator GetEnumerator();

        [Description("Find the index of the first matching String in the " + nameof(ListString) + ", search from start to " +
                "end, optionally specifying an initial index to search from and a count of Strings to search. The " + nameof(ListString) + 
                " must be sorted alphabetically. If no matching String is found, returns a negative number that is the bitwise " +
                "complement of the index of the next String that is larger than the supplied String or, if there is no larger String, " +
                "the bitwise complement of Count. Indexes are zero-based.")]
        [DispId(119)]
        int BinarySearch(string item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue);

        [Description("Remove the first occurrence of the supplied String from the " + nameof(ListString) + ", searching from " +
                "start to end. Returns True if a String was removed, False if the String does not exist in the " + nameof(ListString) + ".")]
        [DispId(120)]
        bool Remove(string item, bool ignoreCase = false);

        [Description("Remove the String at the given index (zero-based) of the " + nameof(ListString) + ".")]
        [DispId(121)]
        void RemoveAt(int index);

        [Description("Remove a range of Strings from the " + nameof(ListString) + " (index is zero-based).")]
        [DispId(122)]
        void RemoveRange(int index, int count);

        [Description("Remove all occurances of all Strings in the " + nameof(ListString) + " that match the supplied Array of " +
                "Strings (the Array must have a 0 lower bound and can be either static or dynamic). Returns the number of Strings " +
                "removed from the " + nameof(ListString) + ".")]
        [DispId(123)]
        int RemoveArray(ref string[] strings, bool ignoreCase = false);

        [Description("Add a " + nameof(ListString) + " to this " + nameof(ListString) + ".")]
        [DispId(124)]
        void AddList(ListString list);

        [Description("Insert a " + nameof(ListString) + " into any valid index (zero-based) of this " + nameof(ListString) + ".")]
        [DispId(125)]
        void InsertList(int index, ListString list);

        [Description("Get a " + nameof(ListString) + " of the items at the specified indeces of this " + nameof(ListString) + ".")]
        [DispId(126)]
        ListString GetList(int index, int count);

        [Description("Remove all occurances of all Strings in this " + nameof(ListString) + " that match a String in the " +
                "supplied " + nameof(ListString) + ". Returns the number of Strings removed from the " + nameof(ListString) + ".")]
        [DispId(127)]
        int RemoveList(ListString list, bool ignoreCase = false);

        [Description("Remove all occurrences of the supplied String from the " + nameof(ListString) + ". Returns the number " +
        "of Strings removed.")]
        [DispId(128)]
        int RemoveAll(string item, bool ignoreCase = false);

        [Description("Insert an Array into any valid index (zero-based) of this " + nameof(ListString) + ". The Array must " +
                "have a 0 lower bound, the Array may be either static or dynamic.")]
        [DispId(129)]
        void InsertArray(int index, ref string[] items);
    }

    [Description("Container for a collection of Strings.")]
    [ProgId("VbaList.ListString")]
    [Guid("87511C9E-3CCB-4271-AF45-4A56EF8C89AC")]
    [ClassInterface(ClassInterfaceType.None)] // ie no auto-generated interface, we're providing a specific interface
    public class ListString : IListString {

        List<string> list;

        public ListString() {
            list = new List<string>();
        }


        public void Add(string item) {
            list.Add(item);
        }


        public string Get(int index) {
            return list[index];
        }


        public int Count() {
            return list.Count;
        }


        public void Sort(bool ascending = true, bool ignoreCase = false) {
            if (ascending) {
                if (ignoreCase) {
                    list.Sort();
                } else {
                    list.Sort(StringComparer.Ordinal);
                }
            } else {
                if (ignoreCase) {
                    list = list.OrderByDescending(s => s).ToList();
                } else {
                    list = list.OrderByDescending(s => s, StringComparer.Ordinal).ToList();
                }
            }
        }


        public void Reverse() {
            list.Reverse();
        }


        public void MakeUnique(bool ignoreCase = false) {
            if (ignoreCase) {
                list = list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            } else {
                list = list.Distinct(StringComparer.Ordinal).ToList();
            }
        }


        public ListIterator GetIterator() {
            return new ListIterator(list);
        }


        public string GetAsText(string sep = ", ", string lastSep = " and ", string ifNone = "", string surround = "") {
            if (list.Count == 0) {
                return ifNone;
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < list.Count; i++) {
                if (i == 0 || i == list.Count) {
                    sb.Append(surround).Append(list[i]).Append(surround);
                } else if (i == list.Count - 1) {
                    sb.Append(lastSep).Append(surround).Append(list[i]).Append(surround);
                } else {
                    sb.Append(sep).Append(surround).Append(list[i]).Append(surround);
                }
            }
            return sb.ToString();
        }


        public string GetAsBullettedList(string bullet = "\u2022 ", string sep = "", string lastSep = "", string ifNone = "",
                string surround = "") {
            if (list.Count == 0) {
                return ifNone;
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < list.Count; i++) {
                if (i == list.Count - 1) {
                    sb.Append(bullet).Append(surround).Append(list[i]).Append(surround);
                } else if (i == list.Count - 2) {
                    sb.Append(bullet).Append(surround).Append(list[i]).Append(surround)
                            .Append(lastSep).Append("\n");
                } else {
                    sb.Append(bullet).Append(surround).Append(list[i]).Append(surround)
                            .Append(sep).Append("\n");
                }
            }
            return sb.ToString();
        }



        public void AddArray(ref string[] strings) {
            list.AddRange(strings);
        }


        public string[] ToArray() {
            return list.ToArray();
        }


        public void Insert(int index, string item) {
            list.Insert(index, item);
        }


        public int IndexOf(string item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue) {
            if (ignoreCase) {
                int startFrom = index == int.MinValue ? 0 : index;
                int endAt = count == int.MinValue ? list.Count : (startFrom + count);
                if (startFrom < 0 || startFrom > list.Count - 1 || endAt > list.Count) {
                    // force an Exception, using the built-in C# exception text, if inputs are out of range
#pragma warning disable IDE0059 // Unnecessary assignment of a value
                    int temp = list.IndexOf(item, startFrom, endAt);
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                }
                for (int i = startFrom; i < endAt; i++) {
                    if (list[i].Equals(item, StringComparison.OrdinalIgnoreCase)) {
                        return i;
                    }
                }
                return -1;
            } else {
                if (count == int.MinValue) {
                    if (index == int.MinValue) {
                        return list.IndexOf(item);
                    } else {
                        return list.IndexOf(item, index);
                    }
                } else {
                    if (index == int.MinValue) {
                        return list.IndexOf(item, 0, count);
                    } else {
                        return list.IndexOf(item, index, count);
                    }

                }
            }
        }


        public int LastIndexOf(string item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue) {
            if (ignoreCase) {
                int startFrom = index == int.MinValue ? list.Count - 1 : index;
                int endAt = count == int.MinValue ? 0 : (startFrom - (count - 1));
                if (startFrom < 0 || startFrom > list.Count - 1 || endAt < 0) {
                    // force an Exception, using the built-in C# exception text, if inputs are out of range
#pragma warning disable IDE0059 // Unnecessary assignment of a value
                    int temp = list.LastIndexOf(item, startFrom, endAt);
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                }
                for (int i = startFrom; i >= endAt; i--) {
                    if (list[i].Equals(item, StringComparison.OrdinalIgnoreCase)) {
                        return i;
                    }
                }
                return -1;
            } else {
                if (count == int.MinValue) {
                    if (index == int.MinValue) {
                        return list.LastIndexOf(item);
                    } else {
                        return list.LastIndexOf(item, index);
                    }
                } else {
                    if (index == int.MinValue) {
                        return list.LastIndexOf(item, list.Count - 1, count);
                    } else {
                        return list.LastIndexOf(item, index, count);
                    }

                }
            }
        }


        public void Clear() {
            list.Clear();
        }


        // for this and all methods that take a "bool ignoreCase" parameter, can also pass either VbCompareMethod.vbBinaryCompare or
        // VbCompareMethod.vbTextCompare
        public bool Contains(string text, bool ignoreCase = false) {
            if (ignoreCase) {
                return list.Any(s => s.Equals(text, StringComparison.OrdinalIgnoreCase));
            } else {
                return list.Contains(text);
            }
        }


        public string GetAsNumberedList(ListType listType, string append = ". ", string sep = "", string lastSep = "", string ifNone = "",
                string surround = "") {
            if (list.Count == 0) {
                return ifNone;
            }
            StringBuilder sb = new StringBuilder();
            string number;
            if (list.Count > 26) {
                listType = ListType.DIGITS;
            }
            for (int i = 0; i < list.Count; i++) {
                if (listType == ListType.DIGITS) {
                    number = (i + 1).ToString();
                } else {
                    number = ((char)(65 + i)).ToString();
                    if (listType == ListType.LOWERCASE) {
                        number = number.ToLower();
                    }
                }
                if (i == list.Count - 1) {
                    sb.Append(number).Append(append).Append(surround).Append(list[i]).Append(surround);
                } else if (i == list.Count - 2) {
                    sb.Append(number).Append(append).Append(surround).Append(list[i])
                            .Append(surround).Append(lastSep).Append("\n");
                } else {
                    sb.Append(number).Append(append).Append(surround).Append(list[i])
                            .Append(surround).Append(sep).Append("\n");
                }
            }
            return sb.ToString();
        }


        // has to be present (and public) to allow use of For ... Each
        public IEnumerator GetEnumerator() {
            return ((IEnumerable)list).GetEnumerator();
        }


        public int BinarySearch(string item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue) {
            if (index == int.MinValue) {
                index = 0;
            }
            if (count == int.MinValue) {
                count = list.Count - index;
            }
            return list.BinarySearch(index, count, item, new MyStringComparer(ignoreCase));
        }


        internal sealed class MyStringComparer : IComparer<string> {

            readonly bool ignoreCase;

            public MyStringComparer(bool ignoreCase) {
                this.ignoreCase = ignoreCase;
            }

            public int Compare(string s1, string s2) {
                if (ignoreCase) {
                    return string.Compare(s1, s2, StringComparison.OrdinalIgnoreCase);
                } else {
                    return string.Compare(s1, s2, StringComparison.Ordinal);
                }
            }
        }


        public bool Remove(string item, bool ignoreCase = false) {
            if (ignoreCase) {
                int index = list.FindIndex(s => s.Equals(item, StringComparison.OrdinalIgnoreCase));
                if (index == -1) {
                    return false;
                } else {
                    list.RemoveAt(index);
                    return true;
                }
            } else {
                return list.Remove(item);
            }
        }


        public void RemoveAt(int index) {
            list.RemoveAt(index);
        }


        public void RemoveRange(int index, int count) {
            list.RemoveRange(index, count);
        }


        public int RemoveArray(ref string[] strings, bool ignoreCase = false) {
            List<string> listToRemove = new List<string>(strings);
            if (ignoreCase) {
                return list.RemoveAll((s) => listToRemove.Contains(s, StringComparer.OrdinalIgnoreCase));
            } else {
                return list.RemoveAll((s) => listToRemove.Contains(s, StringComparer.Ordinal));
            }
        }


        public void AddList(ListString list) {
            if (list.Count() == 0) {
                return;
            }
            this.list.AddRange(list.ToArray());
        }


        public void InsertList(int index, ListString list) {
            if (list.Count() == 0) {
                return;
            }
            this.list.InsertRange(index, list.ToArray());
        }


        public ListString GetList(int index, int count) {
            ListString listString = new ListString();
            listString.list.AddRange(list.GetRange(index, count));
            return listString;
        }


        public int RemoveList(ListString list, bool ignoreCase = false) {
            if (ignoreCase) {
                return this.list.RemoveAll(s => list.list.Contains(s, StringComparer.OrdinalIgnoreCase));
            } else {
                return this.list.RemoveAll(s => list.list.Contains(s, StringComparer.Ordinal));
            }
        }


        public int RemoveAll(string item, bool ignoreCase = false) {
            if (ignoreCase) {
                return list.RemoveAll(s => s.Equals(item, StringComparison.OrdinalIgnoreCase));
            } else {
                return list.RemoveAll(s => s.Equals(item, StringComparison.Ordinal));
            }
        }


        public void InsertArray(int index, ref string[] items) {
            if (items.Length == 0) {
                return;
            }
            list.InsertRange(index, items);
        }


    }
}
