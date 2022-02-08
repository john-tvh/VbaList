using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;


namespace VbaList {


    [Description("Implement '" + nameof(ISortable) + "' in your Class to allow the items in a " + nameof(ListObject) + 
            " to be sorted by one or more values.")]
    [Guid("B4A40C0B-1E6E-425E-AE12-A361AA25AFD3")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // default - allow early and late binding
    public interface ISortable {

        [Description("Allow the items in a " + nameof(ListObject) + " to be sorted by one or more values. The 'compareType' " +
                "parameter is to allow your code to define the basis for the sort. From this procedure, your code must return a Long " +
                "that is to be used to order the items, where: 1 (or greater than zero): this item is greater than 'otherObject'; " +
                "-1 (or less than zero): this item is less than 'otherObject'; 0: the items are equal.")]
        [DispId(502)]
        int CompareTo(object otherObject, int compareType, bool ascending1, bool ignoreCase1, bool ascending2, bool ignoreCase2, 
                bool ascending3, bool ignoreCase3);
    }


    [Description("Implement '" + nameof(IEquals) + "' in your Class to allow the items in a " + nameof(ListObject) + 
            " to be compared for equality.")]
    [Guid("A44FC7B7-A61F-4515-9FDF-7352B47315E0")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // default - allow early and late binding
    public interface IEquals {

        [Description("Allow the items in a " + nameof(ListObject) + " to be compared for equality. From this procedure, your " +
                "code must return a Boolean, where: True: the items are equal; False: the items are non-equal.")]
        [DispId(503123)]
        bool Equals(object otherObject, bool ignoreCase);
    }


    [Description("Implement '" + nameof(IToString) + "' in your Class to allow the items in a " + nameof(ListObject) + 
            " to be returned as String representations.")]
    [Guid("C340D268-6B21-4E71-B276-02D4AA94A298")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // default - allow early and late binding
    public interface IToString {

        [Description("Allow the items in a " + nameof(ListObject) + " to be returned as String representations. From this " +
                "procedure, your code must return a String.")]
        [DispId(504)]
        string ToString();
    }


    [Guid("71127A76-EF93-4002-BACE-DAF99CAD09B9")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // default - allow early and late binding
    public interface IListObject {

        [Description("Add an item to the end of the " + nameof(ListObject) + ".")]
        [DispId(401)]
        void Add(object item);

        [Description("Get the item at a specific index of the " + nameof(ListObject) + ".")]
        [DispId(402)]
        object Get(int index);

        [Description("Get the number of items held in the " + nameof(ListObject) + ".")]
        [DispId(403)]
        int Count();

        [Description("Sort the items in the " + nameof(ListObject) + " by one or more criteria. The Class that defines the " +
                "items must implement the " + nameof(ISortable) + " interface. The 'compareType' and various 'ascending' and 'ignoreCase' " +
                "parameters are passed through to your implementation of " + nameof(ISortable) + "_" + nameof(ISortable.CompareTo) + 
                " to allow your code to define the basis for the sort.")]
        [DispId(405)]
        void Sort(int compareType, bool ascending1 = true, bool ignoreCase1 = false, 
                bool ascending2 = true, bool ignoreCase2 = false, 
                bool ascending3 = true, bool ignoreCase3 = false);

        [Description("Sort the items in the " + nameof(ListObject) + " by a Property of that item.")]
        [DispId(406)]
        void SortByProperty(string propertyName, bool ascending, bool ignoreCase = false);

        [Description("Sort the items in the " + nameof(ListObject) + " by a Method (that has a return value) of that item.")]
        [DispId(407)]
        void SortByMethod(string methodName, bool ascending, bool ignoreCase = false);

        [Description("Reverse the order of items in the " + nameof(ListObject) + ".")]
        [DispId(408)]
        void Reverse();

        [Description("Make every item in the " + nameof(ListObject) + " unique by removing duplicates. The Class that " +
                "defines the items must implement the " + nameof(VbaList) + "." + nameof(IEquals) + " interface.")]
        [DispId(409)]
        void MakeUnique(bool ignoreCase = false);

        [Description("Get a " + nameof(ListIterator) + " to iterate over the " + nameof(ListObject) + ".")]
        [DispId(410)]
        ListIterator GetIterator();

        [Description("Get all items in the " + nameof(ListObject) + " as a single line of text. The Class that defines the " +
                "items must implement the " + nameof(IToString) + " interface.")]
        [DispId(411)]
        string GetAsText(string sep = ", ", string lastSep = " and ", string ifNone = "none", string surround = "");

        [Description("Get all items in the " + nameof(ListObject) + " as a bulleted list.")]
        [DispId(412)]
        string GetAsBullettedList(string bullet = "\u2022 ", string sep = "", string lastSep = "", string ifNone = "",
                string surround = "");

        [Description("Get all items in the " + nameof(ListObject) + " as a numbered list.")]
        [DispId(413)]
        string GetAsNumberedList(ListType listType, string append = ". ", string sep = "", string lastSep = "", string ifNone = "",
        string surround = "");

        [Description("Add an Array of items to the end of the " + nameof(ListObject) + ". The Array must have a Variant " +
                "data type and must have a 0 lower bound, the Array can be either static or dynamic.")]
        [DispId(414)]
        void AddArray(ref object[] items);

        [Description("Return the " + nameof(ListObject) + " as an Array (can either be assigned to a dynamic Array of " +
                "Variants or to a Variant.")]
        [DispId(415)]
        object[] ToArray();

        [Description("Insert an item into any valid index (zero-based) of the " + nameof(ListObject) + ".")]
        [DispId(416)]
        void Insert(int index, object item);

        [Description("Find the index of the first matching item in the " + nameof(ListObject) + ", search from start to " +
                "end, optionally specifying an initial index to search from and a count of Objects to search. Indexes are zero-based. " +
                "The Class that defines the items must implement the " + nameof(IEquals) + " interface.")]
        [DispId(417)]
        int IndexOf(object item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue);

        [Description("Clear all items from the " + nameof(ListObject) + ".")]
        [DispId(418)]
        void Clear();

        [Description("Find the index of the last matching item in the " + nameof(ListObject) + ", search from start to end, " +
                "optionally specifying an initial index to search from and a count of Objects to search. Indexes are zero-based. The " +
                "Class that defines the items must implement the " + nameof(IEquals) + " interface.")]
        [DispId(419)]
        int LastIndexOf(object item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue);

        [Description("Does the " + nameof(ListObject) + " contain the given item. The Class that defines the items must " +
                "implement the " + nameof(IEquals) + " interface.")]
        [DispId(420)]
        bool Contains(object item, bool ignoreCase = false);

        [Description("Supports use of \"For Each ... In\" (do not use this method directly).")]
        [DispId(-4)]
        IEnumerator GetEnumerator();

        [Description("Remove the first occurrence of the supplied item from the " + nameof(ListObject) + " searching from " +
                "start to end. Returns True if an item was removed, False if the item does not exist in the " + nameof(ListObject) +
                ". The Class that defines the items must implement the " + nameof(IEquals) + " interface.")]
        [DispId(421)]
        bool Remove(object item, bool ignoreCase = false);

        [Description("Remove the item at the given index (zero-based) of the " + nameof(ListObject) + ".")]
        [DispId(422)]
        void RemoveAt(int index);

        [Description("Remove a range of items from the " + nameof(ListObject) + " (index is zero-based).")]
        [DispId(423)]
        void RemoveRange(int index, int count);

        [Description("Remove all occurances of all items in the " + nameof(ListObject) + " that match the supplied Array of " +
                "items (must be an Array of Variants; the Array can be either static or dynamic; it must have a 0 lower bound). " +
                "Returns the number of items removed from the " + nameof(ListObject) + ". The Class that defines the items must " +
                "implement the " + nameof(IEquals) + " interface.")]
        [DispId(424)]
        int RemoveArray(ref object[] items, bool ignoreCase = false);

        [Description("Add a " + nameof(ListObject) + " to this " + nameof(ListObject) + ".")]
        [DispId(425)]
        void AddList(ListObject list);

        [Description("Insert a " + nameof(ListObject) + " into any valid index (zero-based) of this " + nameof(ListObject) + ".")]
        [DispId(426)]
        void InsertList(int index, ListObject list);

        [Description("Get a " + nameof(ListObject) + " of the items at the specified indeces of this " + nameof(ListObject) + ".")]
        [DispId(427)]
        ListObject GetList(int index, int count);

        [Description("Remove all occurances of all items in this " + nameof(ListObject) + " that match an item in the " +
                "supplied " + nameof(ListObject) + ". Returns the number of items removed from the " + nameof(ListObject) + ". The " +
                "Class that defines the items must implement the " + nameof(IEquals) + " interface.")]
        [DispId(428)]
        int RemoveList(ListObject list, bool ignoreCase = false);

        [Description("Remove all occurrences of the supplied item from the " + nameof(ListObject) + ". Returns the number " +
                "of items removed. The Class that defines the items must implement the " + nameof(IEquals) + " interface.")]
        [DispId(429)]
        int RemoveAll(object item, bool ignoreCase = false);

        [Description("Insert an Array into any valid index (zero-based) of this " + nameof(ListObject) + ". The Array must " +
                "have a Variant data type and must have a 0 lower bound, the Array may be either static or dynamic.")]
        [DispId(430)]
        void InsertArray(int index, ref object[] items);
    }

    
    [Description("Container for a collection of Objects. It is possible to add multiple types of Objects to a single " + 
            nameof(ListObject) + ", however, procedures that use the " + nameof(IEquals) + " and " + nameof(ISortable) + " interfaces " +
            "may not function correctly unless every type of Object contained in the " + nameof(ListObject) + " implements the " +
            "interface, plus the " + nameof(SortByMethod) + " and " + nameof(SortByProperty) + " methods will fail unless every type " +
            "implements the Property/Method with a return value that can be compared.")]
    [ProgId("VbaList.ListObject")] // must be format foo.bar ... my convention: namespace.class
    [Guid("F54361DC-EABB-47E8-8B30-612D2C34886A")]
    [ClassInterface(ClassInterfaceType.None)] // ie no auto-generated interface, we're providing a specific interface
    public class ListObject : IListObject {

        List<object> list;

        public ListObject() {
            list = new List<object>();
        }


        public void Add(object item) {
            list.Add(item);
        }


        public object Get(int index) {
            return list[index];
        }


        public int Count() {
            return list.Count;
        }


        public void Sort(int compareType, bool ascending1 = true, bool ignoreCase1 = false, 
                bool ascending2 = true, bool ignoreCase2 = false, 
                bool ascending3 = true, bool ignoreCase3 = false) {
            if (list.Count <= 1) {
                return;
            }
            try {
                list.Sort(new SortComparerForListObject(compareType, ascending1, ignoreCase1, 
                        ascending2, ignoreCase2, ascending3, ignoreCase3));
            } catch (Exception e) {
                if (e.InnerException != null && e.InnerException.GetType().Equals(typeof(COMException))) {
                    // the VBA code raised an error
                    throw new InvalidOperationException($"An error \"{e.InnerException.Message}\" was raised from the VBA " +
                            $"implementation of the {nameof(ISortable)}.{nameof(ISortable.CompareTo)}() method.");
                } else {
                    throw new InvalidOperationException($"Failed to perform the sort. Please ensure the {nameof(ListObject)} " +
                                $"contains a single type of object and that the class the object is created from implements the " +
                                $"{nameof(ISortable)} interface and the {nameof(ISortable)}.{nameof(ISortable.CompareTo)} method.");
                }
            }
        }


        internal sealed class SortComparerForListObject : IComparer<object> {
            readonly int compareType;
            readonly bool ascending1, ascending2, ascending3;
            readonly bool ignoreCase1, ignoreCase2, ignoreCase3;

            public SortComparerForListObject(int compareType, bool ascending1, bool ignoreCase1, 
                    bool ascending2, bool ignoreCase2, 
                    bool ascending3, bool ignoreCase3) {
                this.compareType = compareType;
                this.ascending1 = ascending1;
                this.ignoreCase1 = ignoreCase1;
                this.ascending2 = ascending2;
                this.ignoreCase2 = ignoreCase2;
                this.ascending3 = ascending3;
                this.ignoreCase3 = ignoreCase3;
            }

            public int Compare(object o1, object o2) {
                return ((ISortable)o1).CompareTo((ISortable)o2, compareType, 
                        ascending1, ignoreCase1,
                        ascending2, ignoreCase2,
                        ascending3, ignoreCase3);
            }
        }


        public void SortByProperty(string propertyName, bool ascending, bool ignoreCase = false) {
            try {
                list.Sort(new PropertyComparer(propertyName, ascending, ignoreCase));
            } catch (InvalidOperationException) {
                throw new InvalidOperationException($"Could not access the '{propertyName}' property");
            }
        }


        internal sealed class PropertyComparer : IComparer<object> {
            readonly string propertyName;
            readonly bool ascending;
            readonly bool ignoreCase;

            public PropertyComparer(string propertyName, bool ascending, bool ignoreCase) {
                this.propertyName = propertyName;
                this.ascending = ascending;
                this.ignoreCase = ignoreCase;
            }

            public int Compare(object o1, object o2) {
                if (propertyName.Length == 0) {
                    throw new ArgumentException($"No property specified");
                }
                try {
                    string propNameLocal = propertyName + 
                            (propertyName.EndsWith(".", StringComparison.OrdinalIgnoreCase) ? "" : ".");
                    while (propNameLocal.Contains(".") && propNameLocal.Length > 1) {
                        string prop1 = propNameLocal.Substring(0, propNameLocal.IndexOf("."));
                        propNameLocal = propNameLocal.Substring(propNameLocal.IndexOf(".") + 1);
                        o1 = o1.GetType().InvokeMember(prop1, BindingFlags.GetProperty, null, o1, null);
                        o2 = o2.GetType().InvokeMember(prop1, BindingFlags.GetProperty, null, o2, null);
                    }
                } catch (COMException) {
                    throw new InvalidOperationException($"Could not access the '{propertyName}' property");
                }
                return CompareTheReplies(propertyName, "property", o1, o2, ascending, 
                        ignoreCase);
            }
        }


        public void SortByMethod(string methodName, bool ascending, bool ignoreCase = false) {
            try {
                list.Sort(new MethodComparer(methodName, ascending, ignoreCase));
            } catch (InvalidOperationException) {
                throw new InvalidOperationException($"Could not access the '{methodName}' method");
            }
        }


        internal sealed class MethodComparer : IComparer<object> {
            readonly string methodName;
            readonly bool ascending;
            readonly bool ignoreCase;

            public MethodComparer(string methodName, bool ascending, bool ignoreCase) {
                this.methodName = methodName;
                this.ascending = ascending;
                this.ignoreCase = ignoreCase;
            }

            public int Compare(object o1, object o2) {
                object reply1;
                object reply2;
                try {
                    reply1 = o1.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, o1, null);
                    reply2 = o2.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, o2, null);
                } catch (COMException) {
                    throw new InvalidOperationException($"Could not access the '{methodName}' method");
                }
                return CompareTheReplies(methodName, "method", reply1, reply2, ascending, 
                        ignoreCase);
            }
        }


        private static int CompareTheReplies(string memberName, string memberType, object reply1, object reply2,
                bool ascending, bool ignoreCase) {
            if (reply1.GetType() == typeof(string)) {
                // VBA String
                string s1 = (string)reply1;
                string s2 = (string)reply2;
                if (ascending) {
                    if (ignoreCase) {
                        return string.Compare(s1, s2, StringComparison.OrdinalIgnoreCase);
                    } else {
                        return string.Compare(s1, s2, StringComparison.Ordinal);
                    }
                } else {
                    if (ignoreCase) {
                        return string.Compare(s2, s1, StringComparison.OrdinalIgnoreCase);
                    } else {
                        return string.Compare(s2, s1, StringComparison.Ordinal);
                    }
                }
            } else if (reply1.GetType() == typeof(int)) {
                // VBA Long
                int n1 = (int)reply1;
                int n2 = (int)reply2;
                if (n1 == n2) {
                    return 0;
                } else if (ascending) {
                    return n1 > n2 ? 1 : -1;
                } else {
                    return n1 < n2 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(short)) {
                // VBA Integer
                short n1 = (short)reply1;
                short n2 = (short)reply2;
                if (n1 == n2) {
                    return 0;
                } else if (ascending) {
                    return n1 > n2 ? 1 : -1;
                } else {
                    return n1 < n2 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(byte)) {
                // VBA Byte
                byte n1 = (byte)reply1;
                byte n2 = (byte)reply2;
                if (n1 == n2) {
                    return 0;
                } else if (ascending) {
                    return n1 > n2 ? 1 : -1;
                } else {
                    return n1 < n2 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(double)) {
                // VBA Double
                double n1 = (double)reply1;
                double n2 = (double)reply2;
                if (n1 == n2) {
                    return 0;
                } else if (ascending) {
                    return n1 > n2 ? 1 : -1;
                } else {
                    return n1 < n2 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(float)) {
                // VBA Single
                float n1 = (float)reply1;
                float n2 = (float)reply2;
                if (n1 == n2) {
                    return 0;
                } else if (ascending) {
                    return n1 > n2 ? 1 : -1;
                } else {
                    return n1 < n2 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(long)) {
                // VBA Long
                long n1 = (long)reply1;
                long n2 = (long)reply2;
                if (n1 == n2) {
                    return 0;
                } else if (ascending) {
                    return n1 > n2 ? 1 : -1;
                } else {
                    return n1 < n2 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(bool)) {
                // VBA Boolean
                bool b1 = (bool)reply1;
                bool b2 = (bool)reply2;
                if (b1 == b2) {
                    return 0;
                } else if (ascending) {
                    return b2 ? 1 : -1;
                } else {
                    return b1 ? 1 : -1;
                }
            } else if (reply1.GetType() == typeof(DateTime)) {
                // VBA Date
                DateTime d1 = (DateTime)reply1;
                DateTime d2 = (DateTime)reply2;
                if (ascending) {
                    return d1.CompareTo(d2);
                } else {
                    return d2.CompareTo(d1);
                }
            }
            throw new InvalidOperationException($"Object does not implement an accessible '{memberName}' {memberType} that returns a " +
                    $"String, Long, Integer, Byte, Double, Single, LongLong, Boolean or Date");
        }


        public void Reverse() {
            list.Reverse();
        }


        public void MakeUnique(bool ignoreCase = false) {
            if (list.Count <= 1) {
                return;
            }
            list = list.Distinct(new EqualityComparerForListObject<object>(ignoreCase)).ToList();
        }


        internal sealed class EqualityComparerForListObject<T> : IEqualityComparer<T> {
            readonly bool ignoreCase;

            public EqualityComparerForListObject(bool ignoreCase) {
                this.ignoreCase = ignoreCase;
            }

            public bool Equals(T obj1, T obj2) {
                return GetIEqualsEquals(obj1, obj2, ignoreCase);
            }

            public int GetHashCode(T obj) {
                return 0; // a fixed value such that Equals() will be called
            }
        }


        /// <summary>
        /// Call the VBA .IEquals_Equals() method for 'thisObject' to compare it to 'otherObject'
        /// </summary>
        /// <param name="thisObject">The first item being tested for equality</param>
        /// <param name="otherObject">The second item being tested for equality</param>
        /// <param name="ignoreCase">True to ignore case, False for case sensitive</param>
        /// <returns>A bool returned from the call to the VBA .IEquals_Equals() method</returns>
        private static bool GetIEqualsEquals(object thisObject, object otherObject, bool ignoreCase) {
            try {
                return ((IEquals)thisObject).Equals(otherObject, ignoreCase);
            } catch (Exception e) {
                if (e.GetType().Equals(typeof(COMException))) {
                    // the VBA code raised an error
                    throw new InvalidOperationException($"An error \"{e.Message}\" was raised from the VBA implementation of " +
                            $"the {nameof(IEquals)}.{nameof(IEquals.Equals)}() method.");
                } else {
                    throw new InvalidOperationException($"Failed to perform an equality check. Please ensure the " +
                            $"{nameof(ListObject)} contains a single type of object and that the class the object is created from " +
                            $"implements the {nameof(IEquals)} interface and the {nameof(IEquals)}.{nameof(IEquals.Equals)} method.");
                }
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
                    sb.Append(surround).Append(GetIToStringToString(list[i])).Append(surround);
                } else if (i == list.Count - 1) {
                    sb.Append(lastSep).Append(surround).Append(GetIToStringToString(list[i])).Append(surround);
                } else {
                    sb.Append(sep).Append(surround).Append(GetIToStringToString(list[i])).Append(surround);
                }
            }
            return sb.ToString();
        }


        /// <summary>
        /// Call the VBA .IToString_ToString() method for 'obj'
        /// </summary>
        /// <param name="obj">The item to call .IToString_ToString() on</param>
        /// <returns>The String returned from the call to the VBA .IToString_ToString() method</returns>
        private static string GetIToStringToString(object obj) {
            try {
                return ((IToString)obj).ToString();
            } catch (Exception e) {
                if (e.GetType().Equals(typeof(COMException))) {
                    // the VBA code raised an error
                    throw new InvalidOperationException($"An error \"{e.Message}\" was raised from the VBA implementation of " +
                            $"the {nameof(IToString)}.{nameof(IToString.ToString)}() method.");
                } else {
                    throw new InvalidOperationException($"Failed to get an item as a String. Please ensure the class of objects " +
                            $"in the {nameof(ListObject)} implements the {nameof(IToString)} interface and the {nameof(IToString)}." +
                            $"{nameof(IToString.ToString)} method.");
                }
            }
        }


        public string GetAsBullettedList(string bullet = "\u2022 ", string sep = "", string lastSep = "", string ifNone = "",
                string surround = "") {
            if (list.Count == 0) {
                return ifNone;
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < list.Count; i++) {
                if (i == list.Count - 1) {
                    sb.Append(bullet).Append(surround).Append(GetIToStringToString(list[i])).Append(surround);
                } else if (i == list.Count - 2) {
                    sb.Append(bullet).Append(surround).Append(GetIToStringToString(list[i])).Append(surround)
                            .Append(lastSep).Append("\n");
                } else {
                    sb.Append(bullet).Append(surround).Append(GetIToStringToString(list[i])).Append(surround)
                            .Append(sep).Append("\n");
                }
            }
            return sb.ToString();
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
                    sb.Append(number).Append(append).Append(surround).Append(GetIToStringToString(list[i]))
                            .Append(surround);
                } else if (i == list.Count - 2) {
                    sb.Append(number).Append(append).Append(surround).Append(GetIToStringToString(list[i]))
                            .Append(surround).Append(lastSep).Append("\n");
                } else {
                    sb.Append(number).Append(append).Append(surround).Append(GetIToStringToString(list[i]))
                            .Append(surround).Append(sep).Append("\n");
                }
            }
            return sb.ToString();
        }


        public void AddArray(ref object[] items) {
            if (items.Length == 0) {
                return;
            }
            list.AddRange(items);
        }


        public object[] ToArray() {
            return list.ToArray();
        }


        public void Insert(int index, object item) {
            list.Insert(index, item);
        }


        public int IndexOf(object item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue) {
            int startFrom = index == int.MinValue ? 0 : index;
            int endAt = count == int.MinValue ? list.Count : (startFrom + count);
            if (startFrom < 0 || startFrom > list.Count - 1 || endAt > list.Count) {
                // force an Exception, using the built-in C# exception text, if inputs are out of range
#pragma warning disable IDE0059 // Unnecessary assignment of a value
                int temp = list.IndexOf(item, startFrom, endAt);
#pragma warning restore IDE0059 // Unnecessary assignment of a value
            }
            for (int i = startFrom; i < endAt; i++) {
                if (GetIEqualsEquals(item, list[i], ignoreCase)) {
                    return i;
                }
            }
            return -1;
        }


        public int LastIndexOf(object item, bool ignoreCase = false, int index = int.MinValue, int count = int.MinValue) {
            int startFrom = index == int.MinValue ? list.Count - 1 : index;
            int endAt = count == int.MinValue ? 0 : (startFrom - (count - 1));
            if (startFrom < 0 || startFrom > list.Count - 1 || endAt < 0) {
                // force an Exception, using the built-in C# exception text, if inputs are out of range
#pragma warning disable IDE0059 // Unnecessary assignment of a value
                int temp = list.LastIndexOf(item, startFrom, endAt);
#pragma warning restore IDE0059 // Unnecessary assignment of a value
            }
            for (int i = startFrom; i >= endAt; i--) {
                if (GetIEqualsEquals(item, list[i], ignoreCase)) {
                    return i;
                }
            }
            return -1;
        }


        public void Clear() {
            list.Clear();
        }


        // for this and all methods that take a "bool ignoreCase" parameter, can also pass either VbCompareMethod.vbBinaryCompare or
        // VbCompareMethod.vbTextCompare
        public bool Contains(object item, bool ignoreCase = false) {
            for (int i = 0; i < list.Count; i++) {
                if (GetIEqualsEquals(item, list[i], ignoreCase)) {
                    return true;
                }
            }
            return false;
        }


        // has to be present (and public) to allow use of For ... Each
        public IEnumerator GetEnumerator() {
            return ((IEnumerable)list).GetEnumerator();
        }


        public bool Remove(object item, bool ignoreCase = false) {
            for (int i = 0; i < list.Count; i++) {
                if (GetIEqualsEquals(item, list[i], ignoreCase)) {
                    list.RemoveAt(i);
                    return true;
                }
            }
            return false;
        }


        public void RemoveAt(int index) {
            list.RemoveAt(index);
        }


        public void RemoveRange(int index, int count) {
            list.RemoveRange(index, count);
        }


        public int RemoveArray(ref object[] items, bool ignoreCase = false) {
            int countInitial = list.Count;
            foreach (object itemToRemove in items) {
                list = list.Where(o => !GetIEqualsEquals(itemToRemove, o, ignoreCase)).ToList();
            }
            return countInitial - list.Count;
        }


        public void AddList(ListObject list) {
            if (list.Count() == 0) {
                return;
            }
            this.list.AddRange(list.ToArray());
        }


        public void InsertList(int index, ListObject list) {
            if (list.Count() == 0) {
                return;
            }
            this.list.InsertRange(index, list.ToArray());
        }


        public ListObject GetList(int index, int count) {
            ListObject listObject = new ListObject();
            listObject.list.AddRange(list.GetRange(index, count));
            return listObject;
        }


        public int RemoveList(ListObject list, bool ignoreCase = false) {
            int countInitial = this.list.Count;
            foreach (object itemToRemove in list) {
                this.list = this.list.Where(o => !GetIEqualsEquals(itemToRemove, o, ignoreCase))
                        .ToList();
            }
            return countInitial - this.list.Count;
        }


        public int RemoveAll(object item, bool ignoreCase = false) {
            int countInitial = list.Count;
            for (int i = 0; i < list.Count; i++) {
                if (GetIEqualsEquals(item, list[i], ignoreCase)) {
                    list.RemoveAt(i);
                }
            }
            return countInitial - list.Count;
        }


        public void InsertArray(int index, ref object[] items) {
            if (items.Length == 0) {
                return;
            }
            list.InsertRange(index, items);
        }

    }
}
