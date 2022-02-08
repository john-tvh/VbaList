using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

namespace VbaList {


    [Guid("1C165DC2-F9E6-4C1A-8073-E7C95FBF5C2F")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // default - allow early and late binding
    public interface IListIterator {

        [Description("Does the " + nameof(ListIterator) + " have a next item?")]
        [DispId(201)]
        bool HasNext();

        [Description("Get the " + nameof(ListIterator) + "'s next item.")]
        [DispId(202)]
        object GetNext(); 

        [Description("Get the index number of the " + nameof(ListIterator) + "'s next item.")]
        [DispId(203)]
        int GetNextIndex();

        [Description("Get the index number of the " + nameof(ListIterator) + "'s previous item.")]
        [DispId(204)]
        int GetPreviousIndex();

        [Description("Reset the " + nameof(ListIterator) + ".")]
        [DispId(205)]
        void Reset();

    }

    
    [Description("Iterate through the items in a " + nameof(ListString) + " or " + nameof(ListObject) + " ... use the " +
            nameof(ListObject.GetIterator) + "() method of the " + nameof(ListString) + " or " + nameof(ListObject) + " to get an " +
            "instance of this class.")]
    [ProgId("VbaList.ListIterator")] // must be format foo.bar ... my convention: namespace.class
    [Guid("AB62CF11-ABEA-40E5-8B2E-D1FB6DFF6192")]
    [ClassInterface(ClassInterfaceType.None)] // ie no auto-generated interface, we're providing a specific interface

    public class ListIterator : IListIterator {

        int index;
        readonly List<string> listString = null;
        readonly List<object> listObject = null;


        public ListIterator(List<string> list) {
            Reset();
            listString = list;
        }


        public ListIterator(List<object> list) {
            Reset();
            listObject = list;
        }

        public bool HasNext() {
            if (listString != null) {
                return index < listString.Count() - 1;
            } else {
                return index < listObject.Count() - 1;
            }
        }

        public object GetNext() {
            try {
                if (listString != null) {
                    return listString[++index];
                } else {
                    return listObject[++index];
                }
            } catch {
                throw new InvalidOperationException("No more items to return");
            }
        }

        public int GetPreviousIndex() {
            return index;
        }

        public int GetNextIndex() {
            if (listString != null) {
                return index < listString.Count() - 1 ? index + 1 : -1;
            } else {
                return index < listObject.Count() - 1 ? index + 1 : -1;
            }
        }

        public void Reset() {
            index = -1;
        }


    }
}
