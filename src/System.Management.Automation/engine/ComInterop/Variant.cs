// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

#if !SILVERLIGHT // ComObject

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
//using Microsoft.Scripting.Utils;
//using Microsoft.Scripting.Generation;
using Assert = System.Management.Automation.Interpreter.Assert;

namespace System.Management.Automation.ComInterop
{
    /// <summary>
    /// Variant is the basic COM type for late-binding. It can contain any other COM data type.
    /// This type definition precisely matches the unmanaged data layout so that the struct can be passed
    /// to and from COM calls.
    /// </summary>
    [StructLayout(LayoutKind.Explicit)]
    internal struct Variant
    {
#if DEBUG
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2207:InitializeValueTypeStaticFieldsInline")]
        static Variant()
        {
            // Variant size is the size of 4 pointers (16 bytes) on a 32-bit processor,
            // and 3 pointers (24 bytes) on a 64-bit processor.
            int intPtrSize = Marshal.SizeOf(typeof(IntPtr));
            int variantSize = Marshal.SizeOf(typeof(Variant));
            if (intPtrSize == 4)
            {
                Debug.Assert(variantSize == (4 * intPtrSize));
            }
            else
            {
                Debug.Assert(intPtrSize == 8);
                Debug.Assert(variantSize == (3 * intPtrSize));
            }
        }
#endif

        // Most of the data types in the Variant are carried in _typeUnion
        [FieldOffset(0)]
        private TypeUnion _typeUnion;

        // Decimal is the largest data type and it needs to use the space that is normally unused in TypeUnion._wReserved1, etc.
        // Hence, it is declared to completely overlap with TypeUnion. A Decimal does not use the first two bytes, and so
        // TypeUnion._vt can still be used to encode the type.
        [FieldOffset(0)]
        private Decimal _decimal;

        [StructLayout(LayoutKind.Sequential)]
        private struct TypeUnion
        {
            internal ushort _vt;
            internal ushort _wReserved1;
            internal ushort _wReserved2;
            internal ushort _wReserved3;

            internal UnionTypes _unionTypes;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Record
        {
            private IntPtr _record;
            private IntPtr _recordInfo;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1049:TypesThatOwnNativeResourcesShouldBeDisposable")]
        [StructLayout(LayoutKind.Explicit)]
        private struct UnionTypes
        {
            #region Generated Outer Variant union types

            // *** BEGIN GENERATED CODE ***
            // generated by function: gen_UnionTypes from: generate_comdispatch.py

            [FieldOffset(0)]
            internal SByte _i1;
            [FieldOffset(0)]
            internal Int16 _i2;
            [FieldOffset(0)]
            internal Int32 _i4;
            [FieldOffset(0)]
            internal Int64 _i8;
            [FieldOffset(0)]
            internal Byte _ui1;
            [FieldOffset(0)]
            internal UInt16 _ui2;
            [FieldOffset(0)]
            internal UInt32 _ui4;
            [FieldOffset(0)]
            internal UInt64 _ui8;
            [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
            [FieldOffset(0)]
            internal IntPtr _int;
            [FieldOffset(0)]
            internal UIntPtr _uint;
            [FieldOffset(0)]
            internal Int16 _bool;
            [FieldOffset(0)]
            internal Int32 _error;
            [FieldOffset(0)]
            internal Single _r4;
            [FieldOffset(0)]
            internal Double _r8;
            [FieldOffset(0)]
            internal Int64 _cy;
            [FieldOffset(0)]
            internal Double _date;
            [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
            [SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
            [FieldOffset(0)]
            internal IntPtr _bstr;
            [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
            [FieldOffset(0)]
            internal IntPtr _unknown;
            [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
            [FieldOffset(0)]
            internal IntPtr _dispatch;

            // *** END GENERATED CODE ***

            #endregion

            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
            [FieldOffset(0)]
            internal IntPtr _byref;
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
            [FieldOffset(0)]
            internal Record _record;
        }

        public override string ToString()
        {
            return string.Format(CultureInfo.CurrentCulture, "Variant ({0})", VariantType);
        }

        /// <summary>
        /// Primitive types are the basic COM types. It includes valuetypes like ints, but also reference types
        /// like BStrs. It does not include composite types like arrays and user-defined COM types (IUnknown/IDispatch).
        /// </summary>
        internal static bool IsPrimitiveType(VarEnum varEnum)
        {
            switch (varEnum)
            {
                #region Generated Outer Variant IsPrimitiveType

                // *** BEGIN GENERATED CODE ***
                // generated by function: gen_IsPrimitiveType from: generate_comdispatch.py

                case VarEnum.VT_I1:
                case VarEnum.VT_I2:
                case VarEnum.VT_I4:
                case VarEnum.VT_I8:
                case VarEnum.VT_UI1:
                case VarEnum.VT_UI2:
                case VarEnum.VT_UI4:
                case VarEnum.VT_UI8:
                case VarEnum.VT_INT:
                case VarEnum.VT_UINT:
                case VarEnum.VT_BOOL:
                case VarEnum.VT_ERROR:
                case VarEnum.VT_R4:
                case VarEnum.VT_R8:
                case VarEnum.VT_DECIMAL:
                case VarEnum.VT_CY:
                case VarEnum.VT_DATE:
                case VarEnum.VT_BSTR:

                    // *** END GENERATED CODE ***

                    #endregion
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Get the managed object representing the Variant.
        /// </summary>
        /// <returns></returns>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        public object ToObject()
        {
            // Check the simple case upfront
            if (IsEmpty)
            {
                return null;
            }

            switch (VariantType)
            {
                case VarEnum.VT_NULL: return DBNull.Value;

                #region Generated Outer Variant ToObject

                // *** BEGIN GENERATED CODE ***
                // generated by function: gen_ToObject from: generate_comdispatch.py

                case VarEnum.VT_I1: return AsI1;
                case VarEnum.VT_I2: return AsI2;
                case VarEnum.VT_I4: return AsI4;
                case VarEnum.VT_I8: return AsI8;
                case VarEnum.VT_UI1: return AsUi1;
                case VarEnum.VT_UI2: return AsUi2;
                case VarEnum.VT_UI4: return AsUi4;
                case VarEnum.VT_UI8: return AsUi8;
                case VarEnum.VT_INT: return AsInt;
                case VarEnum.VT_UINT: return AsUint;
                case VarEnum.VT_BOOL: return AsBool;
                case VarEnum.VT_ERROR: return AsError;
                case VarEnum.VT_R4: return AsR4;
                case VarEnum.VT_R8: return AsR8;
                case VarEnum.VT_DECIMAL: return AsDecimal;
                case VarEnum.VT_CY: return AsCy;
                case VarEnum.VT_DATE: return AsDate;
                case VarEnum.VT_BSTR: return AsBstr;
                case VarEnum.VT_UNKNOWN: return AsUnknown;
                case VarEnum.VT_DISPATCH: return AsDispatch;
                case VarEnum.VT_VARIANT: return AsVariant;

                // *** END GENERATED CODE ***

                #endregion

                default:
                    return AsVariant;
            }
        }

        /// <summary>
        /// Release any unmanaged memory associated with the Variant
        /// </summary>
        /// <returns></returns>
        public void Clear()
        {
            // We do not need to call OLE32's VariantClear for primitive types or ByRefs
            // to safe ourselves the cost of interop transition.
            // ByRef indicates the memory is not owned by the VARIANT itself while
            // primitive types do not have any resources to free up.
            // Hence, only safearrays, BSTRs, interfaces and user types are
            // handled differently.
            VarEnum vt = VariantType;
            if ((vt & VarEnum.VT_BYREF) != 0)
            {
                VariantType = VarEnum.VT_EMPTY;
            }
            else if (
              ((vt & VarEnum.VT_ARRAY) != 0) ||
              ((vt) == VarEnum.VT_BSTR) ||
              ((vt) == VarEnum.VT_UNKNOWN) ||
              ((vt) == VarEnum.VT_DISPATCH) ||
              ((vt) == VarEnum.VT_RECORD)
              )
            {
                IntPtr variantPtr = UnsafeMethods.ConvertVariantByrefToPtr(ref this);
                NativeMethods.VariantClear(variantPtr);
                Debug.Assert(IsEmpty);
            }
            else
            {
                VariantType = VarEnum.VT_EMPTY;
            }
        }

        public VarEnum VariantType
        {
            get
            {
                return (VarEnum)_typeUnion._vt;
            }

            set
            {
                _typeUnion._vt = (ushort)value;
            }
        }

        internal bool IsEmpty
        {
            get
            {
                return _typeUnion._vt == ((ushort)VarEnum.VT_EMPTY);
            }
        }

        public void SetAsNull()
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = VarEnum.VT_NULL;
        }

        public void SetAsIConvertible(IConvertible value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise

            TypeCode tc = value.GetTypeCode();
            CultureInfo ci = CultureInfo.CurrentCulture;

            switch (tc)
            {
                case TypeCode.Empty: break;
                case TypeCode.Object: AsUnknown = value; break;
                case TypeCode.DBNull: SetAsNull(); break;
                case TypeCode.Boolean: AsBool = value.ToBoolean(ci); break;
                case TypeCode.Char: AsUi2 = value.ToChar(ci); break;
                case TypeCode.SByte: AsI1 = value.ToSByte(ci); break;
                case TypeCode.Byte: AsUi1 = value.ToByte(ci); break;
                case TypeCode.Int16: AsI2 = value.ToInt16(ci); break;
                case TypeCode.UInt16: AsUi2 = value.ToUInt16(ci); break;
                case TypeCode.Int32: AsI4 = value.ToInt32(ci); break;
                case TypeCode.UInt32: AsUi4 = value.ToUInt32(ci); break;
                case TypeCode.Int64: AsI8 = value.ToInt64(ci); break;
                case TypeCode.UInt64: AsI8 = value.ToInt64(ci); break;
                case TypeCode.Single: AsR4 = value.ToSingle(ci); break;
                case TypeCode.Double: AsR8 = value.ToDouble(ci); break;
                case TypeCode.Decimal: AsDecimal = value.ToDecimal(ci); break;
                case TypeCode.DateTime: AsDate = value.ToDateTime(ci); break;
                case TypeCode.String: AsBstr = value.ToString(ci); break;

                default:
                    throw Assert.Unreachable;
            }
        }

        #region Generated Outer Variant accessors

        // *** BEGIN GENERATED CODE ***
        // generated by function: gen_accessors from: generate_comdispatch.py

        // VT_I1
        public SByte AsI1
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_I1);
                return _typeUnion._unionTypes._i1;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_I1;
                _typeUnion._unionTypes._i1 = value;
            }
        }

        public void SetAsByrefI1(ref SByte value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_I1 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertSByteByrefToPtr(ref value);
        }

        // VT_I2
        public Int16 AsI2
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_I2);
                return _typeUnion._unionTypes._i2;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_I2;
                _typeUnion._unionTypes._i2 = value;
            }
        }

        public void SetAsByrefI2(ref Int16 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_I2 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertInt16ByrefToPtr(ref value);
        }

        // VT_I4
        public Int32 AsI4
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_I4);
                return _typeUnion._unionTypes._i4;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_I4;
                _typeUnion._unionTypes._i4 = value;
            }
        }

        public void SetAsByrefI4(ref Int32 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_I4 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertInt32ByrefToPtr(ref value);
        }

        // VT_I8
        public Int64 AsI8
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_I8);
                return _typeUnion._unionTypes._i8;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_I8;
                _typeUnion._unionTypes._i8 = value;
            }
        }

        public void SetAsByrefI8(ref Int64 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_I8 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertInt64ByrefToPtr(ref value);
        }

        // VT_UI1
        public Byte AsUi1
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_UI1);
                return _typeUnion._unionTypes._ui1;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_UI1;
                _typeUnion._unionTypes._ui1 = value;
            }
        }

        public void SetAsByrefUi1(ref Byte value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_UI1 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertByteByrefToPtr(ref value);
        }

        // VT_UI2
        public UInt16 AsUi2
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_UI2);
                return _typeUnion._unionTypes._ui2;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_UI2;
                _typeUnion._unionTypes._ui2 = value;
            }
        }

        public void SetAsByrefUi2(ref UInt16 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_UI2 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertUInt16ByrefToPtr(ref value);
        }

        // VT_UI4
        public UInt32 AsUi4
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_UI4);
                return _typeUnion._unionTypes._ui4;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_UI4;
                _typeUnion._unionTypes._ui4 = value;
            }
        }

        public void SetAsByrefUi4(ref UInt32 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_UI4 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertUInt32ByrefToPtr(ref value);
        }

        // VT_UI8
        public UInt64 AsUi8
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_UI8);
                return _typeUnion._unionTypes._ui8;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_UI8;
                _typeUnion._unionTypes._ui8 = value;
            }
        }

        public void SetAsByrefUi8(ref UInt64 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_UI8 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertUInt64ByrefToPtr(ref value);
        }

        // VT_INT
        public IntPtr AsInt
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_INT);
                return _typeUnion._unionTypes._int;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_INT;
                _typeUnion._unionTypes._int = value;
            }
        }

        public void SetAsByrefInt(ref IntPtr value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_INT | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertIntPtrByrefToPtr(ref value);
        }

        // VT_UINT
        public UIntPtr AsUint
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_UINT);
                return _typeUnion._unionTypes._uint;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_UINT;
                _typeUnion._unionTypes._uint = value;
            }
        }

        public void SetAsByrefUint(ref UIntPtr value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_UINT | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertUIntPtrByrefToPtr(ref value);
        }

        // VT_BOOL
        public bool AsBool
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_BOOL);
                return _typeUnion._unionTypes._bool != 0;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_BOOL;
                _typeUnion._unionTypes._bool = value ? (Int16)(-1) : (Int16)0;
            }
        }

        public void SetAsByrefBool(ref Int16 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_BOOL | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertInt16ByrefToPtr(ref value);
        }

        // VT_ERROR
        public Int32 AsError
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_ERROR);
                return _typeUnion._unionTypes._error;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_ERROR;
                _typeUnion._unionTypes._error = value;
            }
        }

        public void SetAsByrefError(ref Int32 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_ERROR | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertInt32ByrefToPtr(ref value);
        }

        // VT_R4
        public Single AsR4
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_R4);
                return _typeUnion._unionTypes._r4;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_R4;
                _typeUnion._unionTypes._r4 = value;
            }
        }

        public void SetAsByrefR4(ref Single value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_R4 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertSingleByrefToPtr(ref value);
        }

        // VT_R8
        public Double AsR8
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_R8);
                return _typeUnion._unionTypes._r8;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_R8;
                _typeUnion._unionTypes._r8 = value;
            }
        }

        public void SetAsByrefR8(ref Double value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_R8 | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertDoubleByrefToPtr(ref value);
        }

        // VT_DECIMAL
        public Decimal AsDecimal
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_DECIMAL);
                // The first byte of Decimal is unused, but usually set to 0
                Variant v = this;
                v._typeUnion._vt = 0;
                return v._decimal;
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_DECIMAL;
                _decimal = value;
                // _vt overlaps with _decimal, and should be set after setting _decimal
                _typeUnion._vt = (ushort)VarEnum.VT_DECIMAL;
            }
        }

        public void SetAsByrefDecimal(ref Decimal value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_DECIMAL | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertDecimalByrefToPtr(ref value);
        }

        // VT_CY
        public Decimal AsCy
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_CY);
                return Decimal.FromOACurrency(_typeUnion._unionTypes._cy);
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_CY;
                _typeUnion._unionTypes._cy = Decimal.ToOACurrency(value);
            }
        }

        public void SetAsByrefCy(ref Int64 value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_CY | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertInt64ByrefToPtr(ref value);
        }

        // VT_DATE
        public DateTime AsDate
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_DATE);
                return DateTime.FromOADate(_typeUnion._unionTypes._date);
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_DATE;
                _typeUnion._unionTypes._date = value.ToOADate();
            }
        }

        public void SetAsByrefDate(ref Double value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_DATE | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertDoubleByrefToPtr(ref value);
        }

        // VT_BSTR
        public string AsBstr
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_BSTR);
                if (_typeUnion._unionTypes._bstr != IntPtr.Zero)
                {
                    return Marshal.PtrToStringBSTR(_typeUnion._unionTypes._bstr);
                }
                else
                {
                    return null;
                }
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_BSTR;
                if (value != null)
                {
                    Marshal.GetNativeVariantForObject(value, UnsafeMethods.ConvertVariantByrefToPtr(ref this));
                }
            }
        }

        public void SetAsByrefBstr(ref IntPtr value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_BSTR | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertIntPtrByrefToPtr(ref value);
        }

        // VT_UNKNOWN
        public object AsUnknown
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_UNKNOWN);
                if (_typeUnion._unionTypes._dispatch != IntPtr.Zero)
                {
                    return Marshal.GetObjectForIUnknown(_typeUnion._unionTypes._unknown);
                }
                else
                {
                    return null;
                }
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_UNKNOWN;
                if (value != null)
                {
                    _typeUnion._unionTypes._unknown = Marshal.GetIUnknownForObject(value);
                }
            }
        }

        public void SetAsByrefUnknown(ref IntPtr value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_UNKNOWN | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertIntPtrByrefToPtr(ref value);
        }

        // VT_DISPATCH
        public object AsDispatch
        {
            get
            {
                Debug.Assert(VariantType == VarEnum.VT_DISPATCH);
                if (_typeUnion._unionTypes._dispatch != IntPtr.Zero)
                {
                    return Marshal.GetObjectForIUnknown(_typeUnion._unionTypes._dispatch);
                }
                else
                {
                    return null;
                }
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                VariantType = VarEnum.VT_DISPATCH;
                if (value != null)
                {
                    _typeUnion._unionTypes._unknown = Marshal.GetIDispatchForObject(value);
                }
            }
        }

        public void SetAsByrefDispatch(ref IntPtr value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_DISPATCH | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertIntPtrByrefToPtr(ref value);
        }

        // *** END GENERATED CODE ***

        #endregion

        // VT_VARIANT

        public object AsVariant
        {
            get
            {
                return Marshal.GetObjectForNativeVariant(UnsafeMethods.ConvertVariantByrefToPtr(ref this));
            }

            set
            {
                Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
                if (value != null)
                {
                    UnsafeMethods.InitVariantForObject(value, ref this);
                }
            }
        }

        public void SetAsByrefVariant(ref Variant value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            VariantType = (VarEnum.VT_VARIANT | VarEnum.VT_BYREF);
            _typeUnion._unionTypes._byref = UnsafeMethods.ConvertVariantByrefToPtr(ref value);
        }

        // constructs a ByRef variant to pass contents of another variant ByRef.
        public void SetAsByrefVariantIndirect(ref Variant value)
        {
            Debug.Assert(IsEmpty); // The setter can only be called once as VariantClear might be needed otherwise
            Debug.Assert((value.VariantType & VarEnum.VT_BYREF) == 0, "double indirection");

            switch (value.VariantType)
            {
                case VarEnum.VT_EMPTY:
                case VarEnum.VT_NULL:
                    // these cannot combine with VT_BYREF. Should try passing as a variant reference
                    SetAsByrefVariant(ref value);
                    return;
                case VarEnum.VT_RECORD:
                    // VT_RECORD's are weird in that regardless of is the VT_BYREF flag is set or not
                    // they have the same internal representation.
                    _typeUnion._unionTypes._record = value._typeUnion._unionTypes._record;
                    break;
                case VarEnum.VT_DECIMAL:
                    _typeUnion._unionTypes._byref = UnsafeMethods.ConvertDecimalByrefToPtr(ref value._decimal);
                    break;
                default:
                    _typeUnion._unionTypes._byref = UnsafeMethods.ConvertIntPtrByrefToPtr(ref value._typeUnion._unionTypes._byref);
                    break;
            }

            VariantType = (value.VariantType | VarEnum.VT_BYREF);
        }

        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        internal static System.Reflection.PropertyInfo GetAccessor(VarEnum varType)
        {
            switch (varType)
            {
                #region Generated Outer Variant accessors PropertyInfos

                // *** BEGIN GENERATED CODE ***
                // generated by function: gen_accessor_propertyinfo from: generate_comdispatch.py

                case VarEnum.VT_I1: return typeof(Variant).GetProperty("AsI1");
                case VarEnum.VT_I2: return typeof(Variant).GetProperty("AsI2");
                case VarEnum.VT_I4: return typeof(Variant).GetProperty("AsI4");
                case VarEnum.VT_I8: return typeof(Variant).GetProperty("AsI8");
                case VarEnum.VT_UI1: return typeof(Variant).GetProperty("AsUi1");
                case VarEnum.VT_UI2: return typeof(Variant).GetProperty("AsUi2");
                case VarEnum.VT_UI4: return typeof(Variant).GetProperty("AsUi4");
                case VarEnum.VT_UI8: return typeof(Variant).GetProperty("AsUi8");
                case VarEnum.VT_INT: return typeof(Variant).GetProperty("AsInt");
                case VarEnum.VT_UINT: return typeof(Variant).GetProperty("AsUint");
                case VarEnum.VT_BOOL: return typeof(Variant).GetProperty("AsBool");
                case VarEnum.VT_ERROR: return typeof(Variant).GetProperty("AsError");
                case VarEnum.VT_R4: return typeof(Variant).GetProperty("AsR4");
                case VarEnum.VT_R8: return typeof(Variant).GetProperty("AsR8");
                case VarEnum.VT_DECIMAL: return typeof(Variant).GetProperty("AsDecimal");
                case VarEnum.VT_CY: return typeof(Variant).GetProperty("AsCy");
                case VarEnum.VT_DATE: return typeof(Variant).GetProperty("AsDate");
                case VarEnum.VT_BSTR: return typeof(Variant).GetProperty("AsBstr");
                case VarEnum.VT_UNKNOWN: return typeof(Variant).GetProperty("AsUnknown");
                case VarEnum.VT_DISPATCH: return typeof(Variant).GetProperty("AsDispatch");

                // *** END GENERATED CODE ***

                #endregion

                case VarEnum.VT_VARIANT:
                case VarEnum.VT_RECORD:
                case VarEnum.VT_ARRAY:
                    return typeof(Variant).GetProperty("AsVariant");

                default:
                    throw Error.VariantGetAccessorNYI(varType);
            }
        }

        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        internal static System.Reflection.MethodInfo GetByrefSetter(VarEnum varType)
        {
            switch (varType)
            {
                #region Generated Outer Variant byref setter

                // *** BEGIN GENERATED CODE ***
                // generated by function: gen_byref_setters from: generate_comdispatch.py

                case VarEnum.VT_I1: return typeof(Variant).GetMethod("SetAsByrefI1");
                case VarEnum.VT_I2: return typeof(Variant).GetMethod("SetAsByrefI2");
                case VarEnum.VT_I4: return typeof(Variant).GetMethod("SetAsByrefI4");
                case VarEnum.VT_I8: return typeof(Variant).GetMethod("SetAsByrefI8");
                case VarEnum.VT_UI1: return typeof(Variant).GetMethod("SetAsByrefUi1");
                case VarEnum.VT_UI2: return typeof(Variant).GetMethod("SetAsByrefUi2");
                case VarEnum.VT_UI4: return typeof(Variant).GetMethod("SetAsByrefUi4");
                case VarEnum.VT_UI8: return typeof(Variant).GetMethod("SetAsByrefUi8");
                case VarEnum.VT_INT: return typeof(Variant).GetMethod("SetAsByrefInt");
                case VarEnum.VT_UINT: return typeof(Variant).GetMethod("SetAsByrefUint");
                case VarEnum.VT_BOOL: return typeof(Variant).GetMethod("SetAsByrefBool");
                case VarEnum.VT_ERROR: return typeof(Variant).GetMethod("SetAsByrefError");
                case VarEnum.VT_R4: return typeof(Variant).GetMethod("SetAsByrefR4");
                case VarEnum.VT_R8: return typeof(Variant).GetMethod("SetAsByrefR8");
                case VarEnum.VT_DECIMAL: return typeof(Variant).GetMethod("SetAsByrefDecimal");
                case VarEnum.VT_CY: return typeof(Variant).GetMethod("SetAsByrefCy");
                case VarEnum.VT_DATE: return typeof(Variant).GetMethod("SetAsByrefDate");
                case VarEnum.VT_BSTR: return typeof(Variant).GetMethod("SetAsByrefBstr");
                case VarEnum.VT_UNKNOWN: return typeof(Variant).GetMethod("SetAsByrefUnknown");
                case VarEnum.VT_DISPATCH: return typeof(Variant).GetMethod("SetAsByrefDispatch");

                // *** END GENERATED CODE ***

                #endregion

                case VarEnum.VT_VARIANT:
                    return typeof(Variant).GetMethod("SetAsByrefVariant");
                case VarEnum.VT_RECORD:
                case VarEnum.VT_ARRAY:
                    return typeof(Variant).GetMethod("SetAsByrefVariantIndirect");

                default:
                    throw Error.VariantGetAccessorNYI(varType);
            }
        }
    }
}

#endif

