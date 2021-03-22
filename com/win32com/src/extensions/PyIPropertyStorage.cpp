// This file implements the IPropertyStorage Interface and Gateway for Python.
// Generated by makegw.py

#include "stdafx.h"
#include "PythonCOM.h"
#include "PythonCOMServer.h"

#ifndef NO_PYCOM_IPROPERTYSTORAGE
#include "PyIPropertyStorage.h"

// @doc - This file contains autoduck documentation
// ---------------------------------------------------

void PyObject_FreePROPSPECs(PROPSPEC *pFree, ULONG cFree)
{
    if (!pFree)
        return;
    for (ULONG i = 0; i < cFree; i++)
        if (pFree[i].ulKind == PRSPEC_LPWSTR && pFree[i].lpwstr)
            PyWinObject_FreeWCHAR(pFree[i].lpwstr);
    free(pFree);
}

PyObject *PyWinObject_FromPROPIDs(const PROPID *propids, ULONG cpropids)
{
    PyObject *ret = PyTuple_New(cpropids);
    if (ret == NULL)
        return NULL;
    for (ULONG i = 0; i < cpropids; i++) {
        PyObject *item = PyLong_FromUnsignedLong(propids[i]);
        if (item == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyTuple_SET_ITEM(ret, i, item);
    }
    return ret;
}

// @object PROPSPEC|Identifies a property.  Can be either an int property id, or a str/unicode property name.
BOOL PyWinObject_AsPROPSPECs(PyObject *ob, PROPSPEC **ppRet, ULONG *pcRet)
{
    TmpPyObject tuple = PyWinSequence_Tuple(ob, pcRet);
    if (tuple == NULL)
        return FALSE;
    size_t numBytes = sizeof(PROPSPEC) * *pcRet;
    *ppRet = (PROPSPEC *)malloc(numBytes);
    if (*ppRet == NULL) {
        PyErr_NoMemory();
        return FALSE;
    }
    ZeroMemory(*ppRet, numBytes);
    for (DWORD i = 0; i < *pcRet; i++) {
        PyObject *sub = PyTuple_GET_ITEM((PyObject *)tuple, i);
        (*ppRet)[i].propid = PyInt_AsUnsignedLongMask(sub);
        if ((*ppRet)[i].propid != (ULONG)-1 || !PyErr_Occurred())
            (*ppRet)[i].ulKind = PRSPEC_PROPID;
        else {
            PyErr_Clear();
            (*ppRet)[i].lpwstr = NULL;
            if (PyWinObject_AsWCHAR(sub, &(*ppRet)[i].lpwstr))
                (*ppRet)[i].ulKind = PRSPEC_LPWSTR;
            else {
                PyErr_Clear();
                PyErr_SetString(PyExc_TypeError, "PROPSPECs must be a sequence of strings or integers");
                PyObject_FreePROPSPECs(*ppRet, *pcRet);
                *ppRet = NULL;
                return FALSE;
            }
        }
    }
    return TRUE;
}

PyObject *PyWinObject_FromPROPSPECs(const PROPSPEC *propspecs, ULONG cpropspecs)
{
    PyObject *ret = PyTuple_New(cpropspecs);
    if (ret == NULL)
        return NULL;
    PyObject *item;
    for (DWORD i = 0; i < cpropspecs; i++) {
        item = NULL;
        if (propspecs[i].ulKind == PRSPEC_PROPID)
            item = PyLong_FromUnsignedLong(propspecs[i].propid);
        else if (propspecs[i].ulKind == PRSPEC_LPWSTR)
            item = PyWinObject_FromWCHAR(propspecs[i].lpwstr);
        else
            PyErr_SetString(PyExc_NotImplementedError, "Unknown PROPSPEC type");
        if (item == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyTuple_SET_ITEM(ret, i, item);
    }
    return ret;
}

// Generic conversion from VT_VECTOR arrays to list.
template <typename arraytype>
PyObject *VectorToSeq(arraytype *A, ULONG count, PyObject *(*converter)(arraytype))
{
    PyObject *ret = PyList_New(count);
    if (ret == NULL)
        return NULL;
    for (ULONG i = 0; i < count; i++) {
        PyObject *subitem = (*converter)(A[i]);
        if (subitem == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyList_SET_ITEM(ret, i, subitem);
    }
    return ret;
}

// Some converters take a reference, eg PyWinObject_FromLARGE_INTEGER
template <typename arraytype>
PyObject *VectorToSeq(arraytype *A, ULONG count, PyObject *(*converter)(arraytype &))
{
    PyObject *ret = PyList_New(count);
    if (ret == NULL)
        return NULL;
    for (ULONG i = 0; i < count; i++) {
        PyObject *subitem = (*converter)(A[i]);
        if (subitem == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyList_SET_ITEM(ret, i, subitem);
    }
    return ret;
}

// ... and some take a const reference, eg PyWinObject_FromIID
template <typename arraytype>
PyObject *VectorToSeq(arraytype *A, ULONG count, PyObject *(*converter)(const arraytype &))
{
    PyObject *ret = PyList_New(count);
    if (ret == NULL)
        return NULL;
    for (ULONG i = 0; i < count; i++) {
        PyObject *subitem = (*converter)(A[i]);
        if (subitem == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyList_SET_ITEM(ret, i, subitem);
    }
    return ret;
}

// Some helper functions for the Vector conversion template
PyObject *PyWinObject_FromCHAR(CHAR c) { return PyInt_FromLong(c); }

PyObject *PyWinObject_FromUCHAR(UCHAR uc) { return PyInt_FromLong(uc); }

PyObject *PyWinObject_FromSHORT(SHORT s) { return PyInt_FromLong(s); }

PyObject *PyWinObject_FromUSHORT(USHORT us) { return PyInt_FromLong(us); }

PyObject *PyWinObject_FromFLOAT(FLOAT f) { return PyFloat_FromDouble(f); }

PyObject *PyWinObject_FromVT_BSTR(BSTR b) { return PyWinObject_FromBstr(b, FALSE); }

PyObject *PyWinObject_FromVARIANT_BOOL(VARIANT_BOOL b)
{
    PyObject *ret;
    if (b == VARIANT_TRUE)
        ret = Py_True;
    else if (b == VARIANT_FALSE)
        ret = Py_False;
    else {
        ret = NULL;
        PyErr_Format(PyExc_ValueError, "Invalid value for VARIANT_BOOL");
    }
    Py_XINCREF(ret);
    return ret;
}

PyObject *PyObject_FromPROPVARIANT(PROPVARIANT *pVar)
{
    switch (pVar->vt) {
        case VT_EMPTY:
        case VT_NULL:
        case VT_ILLEGAL:
            Py_INCREF(Py_None);
            return Py_None;
        case VT_I1:
            return PyInt_FromLong(pVar->cVal);
        case VT_I1 | VT_VECTOR:
            return VectorToSeq(pVar->cac.pElems, pVar->cac.cElems, PyWinObject_FromCHAR);
        case VT_UI1:
            return PyInt_FromLong(pVar->bVal);
        case VT_UI1 | VT_VECTOR:
            return VectorToSeq(pVar->caub.pElems, pVar->caub.cElems, PyWinObject_FromUCHAR);
        case VT_I2:
            return PyInt_FromLong(pVar->iVal);
        case VT_I2 | VT_VECTOR:
            return VectorToSeq(pVar->cai.pElems, pVar->cai.cElems, PyWinObject_FromSHORT);
        case VT_UI2:
            return PyInt_FromLong(pVar->uiVal);
        case VT_UI2 | VT_VECTOR:
            return VectorToSeq(pVar->caui.pElems, pVar->caui.cElems, PyWinObject_FromUSHORT);
        case VT_I4:
            return PyInt_FromLong(pVar->lVal);
        case VT_I4 | VT_VECTOR:
            return VectorToSeq(pVar->cal.pElems, pVar->cal.cElems, PyInt_FromLong);
        case VT_INT:
            return PyInt_FromLong(pVar->intVal);
        case VT_UI4:
            return PyLong_FromUnsignedLong(pVar->ulVal);
        case VT_UI4 | VT_VECTOR:
            return VectorToSeq(pVar->caul.pElems, pVar->caul.cElems, PyLong_FromUnsignedLong);
        case VT_UINT:
            return PyLong_FromUnsignedLong(pVar->uintVal);
        case VT_I8:
            return PyWinObject_FromLARGE_INTEGER(pVar->hVal);
        case VT_I8 | VT_VECTOR:
            return VectorToSeq<LARGE_INTEGER>(pVar->cah.pElems, pVar->cah.cElems, PyWinObject_FromLARGE_INTEGER);
        case VT_UI8:
            return PyWinObject_FromULARGE_INTEGER(pVar->uhVal);
        case VT_UI8 | VT_VECTOR:
            return VectorToSeq<ULARGE_INTEGER>(pVar->cauh.pElems, pVar->cauh.cElems, PyWinObject_FromULARGE_INTEGER);
        case VT_R4:
            return PyFloat_FromDouble(pVar->fltVal);
        case VT_R4 | VT_VECTOR:
            return VectorToSeq(pVar->caflt.pElems, pVar->caflt.cElems, PyWinObject_FromFLOAT);
        case VT_R8:
            return PyFloat_FromDouble(pVar->dblVal);
        case VT_R8 | VT_VECTOR:
            return VectorToSeq(pVar->cadbl.pElems, pVar->cadbl.cElems, PyFloat_FromDouble);
        case VT_CY:
            return PyObject_FromCurrency(pVar->cyVal);
        case VT_CY | VT_VECTOR:
            return VectorToSeq<CY>(pVar->cacy.pElems, pVar->cacy.cElems, PyObject_FromCurrency);
        case VT_DATE:
            return PyWinObject_FromDATE(pVar->date);
        case VT_DATE | VT_VECTOR:
            return VectorToSeq(pVar->cadate.pElems, pVar->cadate.cElems, PyWinObject_FromDATE);
        case VT_BSTR:
            return PyWinObject_FromBstr(pVar->bstrVal);
        case VT_BSTR | VT_VECTOR:
            return VectorToSeq(pVar->cabstr.pElems, pVar->cabstr.cElems, PyWinObject_FromVT_BSTR);
        case VT_BOOL:
            return PyWinObject_FromVARIANT_BOOL(pVar->boolVal);
        case VT_BOOL | VT_VECTOR:
            return VectorToSeq(pVar->cabool.pElems, pVar->cabool.cElems, PyWinObject_FromVARIANT_BOOL);
        case VT_ERROR:
            return PyInt_FromLong(pVar->scode);
        case VT_ERROR | VT_VECTOR:
            return VectorToSeq(pVar->cascode.pElems, pVar->cascode.cElems, PyInt_FromLong);
        case VT_FILETIME:
            return PyWinObject_FromFILETIME(pVar->filetime);
        case VT_FILETIME | VT_VECTOR:
            return VectorToSeq<FILETIME>(pVar->cafiletime.pElems, pVar->cafiletime.cElems, PyWinObject_FromFILETIME);
        case VT_LPSTR:
            if (pVar->pszVal == NULL) {
                Py_INCREF(Py_None);
                return Py_None;
            }
            return PyWinCoreString_FromString(pVar->pszVal);
        case VT_LPSTR | VT_VECTOR: {
            PyObject *ret = PyList_New(pVar->calpstr.cElems);
            if (ret == NULL)
                return NULL;
            for (ULONG i = 0; i < pVar->calpstr.cElems; i++) {
                PyObject *elem = PyWinCoreString_FromString(pVar->calpstr.pElems[i]);
                if (elem == NULL) {
                    Py_DECREF(ret);
                    return NULL;
                }
                PyList_SET_ITEM(ret, i, elem);
            }
            return ret;
        }
        case VT_LPWSTR:
            return PyWinObject_FromOLECHAR(pVar->pwszVal);
        case VT_LPWSTR | VT_VECTOR: {
            PyObject *ret = PyList_New(pVar->calpwstr.cElems);
            if (ret == NULL)
                return NULL;
            for (ULONG i = 0; i < pVar->calpwstr.cElems; i++) {
                PyObject *elem = PyWinObject_FromWCHAR(pVar->calpwstr.pElems[i]);
                if (elem == NULL) {
                    Py_DECREF(ret);
                    return NULL;
                }
                PyList_SET_ITEM(ret, i, elem);
            }
            return ret;
        }
        case VT_CLSID:
            return PyWinObject_FromIID(*pVar->puuid);
        case VT_CLSID | VT_VECTOR:
            return VectorToSeq<CLSID>(pVar->cauuid.pElems, pVar->cauuid.cElems, PyWinObject_FromIID);
        case VT_STREAM:
        case VT_STREAMED_OBJECT:
            return PyCom_PyObjectFromIUnknown(pVar->pStream, IID_IStream, TRUE);
        case VT_STORAGE:
        case VT_STORED_OBJECT:
            return PyCom_PyObjectFromIUnknown(pVar->pStorage, IID_IStorage, TRUE);
        case VT_VECTOR | VT_VARIANT:
            return PyObject_FromPROPVARIANTs(pVar->capropvar.pElems, pVar->capropvar.cElems);
        case VT_BLOB:
        case VT_BLOB_OBJECT:
            return PyString_FromStringAndSize((const char *)pVar->blob.pBlobData, pVar->blob.cbSize);
            //		case VT_UNKNOWN:
            //			return PyCom_PyObjectFromIUnknown(pVar->punkVal, IID_IUnknown, TRUE);
            //		case VT_DISPATCH:
            //			return PyCom_PyObjectFromIUnknown(pVar->pdispVal, IID_IDispatch, TRUE);

            /*
            // Want to get VT_CF and VT_BLOB working with a test case first!
                    case VT_CF: { // special "clipboard format"
                        // cbSize is the size of the buffer pointed to
                        // by pClipData, plus sizeof(ulClipFmt)
                        // XXX - in that case, shouldn't we pass
                        // pClipData + sizeof(DWORD) to Py_BuildValue??
                        ULONG cb = CBPCLIPDATA(*pVar->pclipdata);
                        return Py_BuildValue("is#",
                                             pVar->pclipdata->ulClipFmt,
                                             pVar->pclipdata->pClipData,
                                             (int)cb);
                        }
            */
        default:
            PyErr_Format(PyExc_TypeError, "Unsupported property type 0x%x", pVar->vt);
            return NULL;
    }
}

PyObject *PyObject_FromPROPVARIANTs(PROPVARIANT *pVars, ULONG cVars)
{
    PyObject *ret = PyTuple_New(cVars);
    if (ret == NULL)
        return NULL;
    for (ULONG i = 0; i < cVars; i++) {
        PyObject *sub = PyObject_FromPROPVARIANT(pVars + i);
        if (sub == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyTuple_SET_ITEM(ret, i, sub);
    }
    return ret;
}

BOOL PyObject_AsPROPVARIANT(PyObject *ob, PROPVARIANT *pVar)
{
    if (ob == Py_None) {
        PropVariantInit(pVar);
    }
    else if (ob == Py_True) {
        pVar->boolVal = -1;
        pVar->vt = VT_BOOL;
    }
    else if (ob == Py_False) {
        pVar->boolVal = 0;
        pVar->vt = VT_BOOL;
    }
    else if (PyLong_Check(ob)) {
        pVar->hVal.QuadPart = PyLong_AsLongLong(ob);
        if (pVar->hVal.QuadPart == -1 && PyErr_Occurred()) {
            // Could still fit in an unsigned long long
            PyErr_Clear();
            pVar->uhVal.QuadPart = PyLong_AsUnsignedLongLong(ob);
            if (pVar->uhVal.QuadPart == -1 && PyErr_Occurred())
                return FALSE;
            pVar->vt = VT_UI8;
        }
        else {
            pVar->vt = VT_I8;
            // Could still fit in a regular long
            if (pVar->hVal.QuadPart >= LONG_MIN && pVar->hVal.QuadPart <= LONG_MAX) {
                pVar->lVal = (long)pVar->hVal.QuadPart;
                pVar->vt = VT_I4;
            }
            // ... or an unsigned long
            else if (pVar->hVal.QuadPart >= 0 && pVar->hVal.QuadPart <= ULONG_MAX) {
                pVar->ulVal = (unsigned long)pVar->hVal.QuadPart;
                pVar->vt = VT_UI4;
            }
        }
#if (PY_VERSION_HEX < 0x03000000)
        // Not needed in Py3k, as PyInt_Check is defined to PyLong_Check
    }
    else if (PyInt_Check(ob)) {
        pVar->lVal = PyInt_AsLong(ob);
        pVar->vt = VT_I4;
#endif
    }
    else if (PyFloat_Check(ob)) {
        pVar->dblVal = PyFloat_AsDouble(ob);
        pVar->vt = VT_R8;
    }
    else if (PyUnicode_Check(ob) || PyString_Check(ob)) {
        PyWinObject_AsBstr(ob, &pVar->bstrVal);
        pVar->vt = VT_BSTR;
    }
    else {
        PyErr_SetString(PyExc_TypeError, "Unsupported object for PROPVARIANT");
        return FALSE;
    }
    return TRUE;
}

void PyObject_FreePROPVARIANTs(PROPVARIANT *pVars, ULONG cVars)
{
    if (pVars) {
        for (ULONG i = 0; i < cVars; i++) PropVariantClear(pVars + i);
        delete[] pVars;
    }
}

BOOL PyObject_AsPROPVARIANTs(PyObject *ob, PROPVARIANT **ppRet, ULONG *pcRet)
{
    BOOL ret = FALSE;
    DWORD len, i;
    PyObject *tuple = PyWinSequence_Tuple(ob, &len);
    if (tuple == NULL)
        return FALSE;

    PROPVARIANT *pRet = new PROPVARIANT[len];
    if (pRet == NULL) {
        PyErr_NoMemory();
        goto cleanup;
    }
    for (i = 0; i < len; i++) PropVariantInit(pRet + i);

    for (i = 0; i < len; i++) {
        PyObject *sub = PyTuple_GET_ITEM(tuple, i);
        if (!PyObject_AsPROPVARIANT(sub, pRet + i))
            goto cleanup;
    }
    ret = TRUE;
cleanup:
    if (ret) {
        *ppRet = pRet;
        *pcRet = len;
    }
    else if (pRet)
        PyObject_FreePROPVARIANTs(pRet, len);
    Py_DECREF(tuple);
    return ret;
}

// Same conversion as PyObject_AsPROPVARIANTs, but PROPVARIANT array is allocated by caller
BOOL PyObject_AsPreallocatedPROPVARIANTs(PyObject *ob, PROPVARIANT *propvars, ULONG cpropvars)
{
    DWORD len, i;
    TmpPyObject tuple = PyWinSequence_Tuple(ob, &len);
    if (tuple == NULL)
        return FALSE;
    if (len != cpropvars) {
        PyErr_SetString(PyExc_ValueError, "Sequence not of required length");
        return FALSE;
    }
    for (i = 0; i < cpropvars; i++) {
        PyObject *sub = PyTuple_GET_ITEM((PyObject *)tuple, i);
        if (!PyObject_AsPROPVARIANT(sub, &propvars[i]))
            for (; i--;) PropVariantClear(&propvars[i]);
        return FALSE;
    }
    return TRUE;
}

BOOL PyObject_AsPROPIDs(PyObject *ob, PROPID **ppRet, ULONG *pcRet)
{
    // PROPID and DWORD are both unsigned long
    return PyWinObject_AsDWORDArray(ob, ppRet, pcRet, FALSE);
}

void PyObject_FreePROPIDs(PROPID *pFree, ULONG)
{
    if (pFree)
        free(pFree);
}

//
// Interface Implementation

PyIPropertyStorage::PyIPropertyStorage(IUnknown *pdisp) : PyIUnknown(pdisp) { ob_type = &type; }

PyIPropertyStorage::~PyIPropertyStorage() {}

/* static */ IPropertyStorage *PyIPropertyStorage::GetI(PyObject *self)
{
    return (IPropertyStorage *)PyIUnknown::GetI(self);
}

// @pymethod (object, ...)|PyIPropertyStorage|ReadMultiple|Reads specified properties from the current property set.
// @rdesc Returned values are automatically converted to an appropriate python type
PyObject *PyIPropertyStorage::ReadMultiple(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    PyObject *props;
    // @pyparm (<o PROPSPEC>, ...)|props||Sequence of property IDs or names.
    if (!PyArg_ParseTuple(args, "O:ReadMultiple", &props))
        return NULL;
    ULONG cProps;
    PROPSPEC *pProps;
    if (!PyWinObject_AsPROPSPECs(props, &pProps, &cProps))
        return NULL;
    PROPVARIANT *pPropVars = new PROPVARIANT[cProps];
    if (pPropVars == NULL) {
        PyObject_FreePROPSPECs(pProps, cProps);
        PyErr_SetString(PyExc_MemoryError, "allocating PROPVARIANTs");
        return NULL;
    }
    ULONG i;
    for (i = 0; i < cProps; i++) PropVariantInit(pPropVars + i);

    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->ReadMultiple(cProps, pProps, pPropVars);
    PY_INTERFACE_POSTCALL;

    PyObject *rc;
    if (FAILED(hr))
        rc = PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    else
        rc = PyObject_FromPROPVARIANTs(pPropVars, cProps);

    // Cleanup the property IDs.
    PyObject_FreePROPSPECs(pProps, cProps);
    // Cleanup the prop variants.
    for (i = 0; i < cProps; i++) {
        PropVariantClear(pPropVars + i);
    }
    delete[] pPropVars;
    return rc;
}

// @pymethod |PyIPropertyStorage|WriteMultiple|Creates or modifies properties in the property set
PyObject *PyIPropertyStorage::WriteMultiple(PyObject *self, PyObject *args)
{
    PyObject *ret = NULL;
    PROPSPEC *pProps = NULL;
    PROPVARIANT *pVals = NULL;
    ULONG cProps, cVals;

    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    PyObject *obProps;
    PyObject *obValues;
    long minId = 2;
    // @pyparm (<o PROPSPEC>, ...)|props||Sequence containing names or integer ids of properties to write
    // @pyparm (<o PROPVARIANT>, ...)|values||The values for the properties.
    // @pyparm int|propidNameFirst|2|Minimum property id to be assigned to new properties specified by name
    if (!PyArg_ParseTuple(args, "OO|l:WriteMultiple", &obProps, &obValues, &minId))
        return NULL;

    if (!PyWinObject_AsPROPSPECs(obProps, &pProps, &cProps))
        goto cleanup;
    if (!PyObject_AsPROPVARIANTs(obValues, &pVals, &cVals))
        goto cleanup;

    if (cProps != cVals) {
        PyErr_SetString(PyExc_ValueError, "The parameters must be sequences of the same size");
        goto cleanup;
    }

    HRESULT hr;
    {
        PY_INTERFACE_PRECALL;
        hr = pIPS->WriteMultiple(cProps, pProps, pVals, minId);
        PY_INTERFACE_POSTCALL;
    }
    if (FAILED(hr))
        PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    else {
        Py_INCREF(Py_None);
        ret = Py_None;
    }

cleanup:
    PyObject_FreePROPSPECs(pProps, cProps);
    PyObject_FreePROPVARIANTs(pVals, cVals);
    return ret;
}

// @pymethod |PyIPropertyStorage|DeleteMultiple|Deletes properties from the property set
PyObject *PyIPropertyStorage::DeleteMultiple(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;

    PyObject *props;
    // @pyparm (<o PROPSPEC>, ...)|props||Sequence containing names or IDs of properties to be deleted
    if (!PyArg_ParseTuple(args, "O:ReadMultiple", &props))
        return NULL;
    ULONG cProps;
    PROPSPEC *pProps;
    if (!PyWinObject_AsPROPSPECs(props, &pProps, &cProps))
        return NULL;

    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->DeleteMultiple(cProps, pProps);
    PY_INTERFACE_POSTCALL;

    PyObject_FreePROPSPECs(pProps, cProps);

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    Py_INCREF(Py_None);
    return Py_None;
}

// @pymethod (str,...)|PyIPropertyStorage|ReadPropertyNames|Retrieves any existing string names for the specified
// property identifiers.
PyObject *PyIPropertyStorage::ReadPropertyNames(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    PyObject *obProps;
    // @pyparm (int, ...)|props||Sequence of ints containing property IDs.
    if (!PyArg_ParseTuple(args, "O:ReadPropertyNames", &obProps))
        return NULL;

    ULONG cProps;
    PROPID *pProps;
    if (!PyObject_AsPROPIDs(obProps, &pProps, &cProps))
        return NULL;

    HRESULT hr;
    LPWSTR *ppStrs = new LPWSTR[cProps];
    if (ppStrs == NULL) {
        PyErr_NoMemory();
        goto cleanup;
    }
    memset(ppStrs, 0, sizeof(LPWSTR) * cProps);
    {
        PY_INTERFACE_PRECALL;
        hr = pIPS->ReadPropertyNames(cProps, pProps, ppStrs);
        PY_INTERFACE_POSTCALL;
    }
    PyObject *rc;
    if (FAILED(hr))
        rc = PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    else {
        rc = PyTuple_New(cProps);
        if (rc == NULL)
            goto cleanup;
        for (ULONG i = 0; i < cProps; i++) {
            PyObject *propname = PyWinObject_FromOLECHAR(ppStrs[i]);
            if (propname == NULL) {
                Py_DECREF(rc);
                rc = NULL;
                goto cleanup;
            }
            PyTuple_SET_ITEM(rc, i, propname);
        }
    }

cleanup:
    if (ppStrs) {
        for (ULONG i = 0; i < cProps; i++)
            if (ppStrs[i])
                CoTaskMemFree(ppStrs[i]);
        delete[] ppStrs;
    }
    PyObject_FreePROPIDs(pProps, cProps);
    return rc;
}

// @pymethod |PyIPropertyStorage|WritePropertyNames|Assigns string names to a specified array of property IDs in the
// current property set.
PyObject *PyIPropertyStorage::WritePropertyNames(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    PyObject *obProps;
    PyObject *obNames;
    // @pyparm (int, ...)|props||Sequence containing the property IDs.
    // @pyparm (string, ...)|names||Equal length sequence of property names.
    if (!PyArg_ParseTuple(args, "OO:WritePropertyNames", &obProps, &obNames))
        return NULL;

    ULONG cProps = 0, cNames = 0;
    PROPID *pProps = NULL;
    LPWSTR *ppStrs = NULL;
    PyObject *rc = NULL;

    if (!PyObject_AsPROPIDs(obProps, &pProps, &cProps))
        return NULL;
    if (!PyWinObject_AsWCHARArray(obNames, &ppStrs, &cNames, FALSE))
        goto done;
    if (cNames != cProps) {
        PyErr_SetString(PyExc_TypeError, "Property names must be a sequence the same size as property ids");
        goto done;
    }

    HRESULT hr;
    {
        PY_INTERFACE_PRECALL;
        hr = pIPS->WritePropertyNames(cProps, pProps, ppStrs);
        PY_INTERFACE_POSTCALL;
    }

    if (FAILED(hr))
        PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    else {
        Py_INCREF(Py_None);
        rc = Py_None;
    }
done:
    PyObject_FreePROPIDs(pProps, cProps);
    PyWinObject_FreeWCHARArray(ppStrs, cNames);
    return rc;
}

// @pymethod |PyIPropertyStorage|DeletePropertyNames|Removes property names from specified properties.
PyObject *PyIPropertyStorage::DeletePropertyNames(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    PyObject *obProps;
    // @pyparm (int, ...)|props||Sequence of ints containing property IDs.
    if (!PyArg_ParseTuple(args, "O:DeletePropertyNames", &obProps))
        return NULL;

    ULONG cProps;
    PROPID *pProps;
    if (!PyObject_AsPROPIDs(obProps, &pProps, &cProps))
        return NULL;
    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->DeletePropertyNames(cProps, pProps);
    PY_INTERFACE_POSTCALL;
    PyObject_FreePROPIDs(pProps, cProps);
    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    Py_INCREF(Py_None);
    return Py_None;
}

// @pymethod |PyIPropertyStorage|Commit|Persists the property set to its base storage
PyObject *PyIPropertyStorage::Commit(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    // @pyparm int|CommitFlags||Combination of storagecon.STGC_* flags
    DWORD grfCommitFlags;
    if (!PyArg_ParseTuple(args, "l:Commit", &grfCommitFlags))
        return NULL;
    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->Commit(grfCommitFlags);
    PY_INTERFACE_POSTCALL;

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    Py_INCREF(Py_None);
    return Py_None;
}

// @pymethod |PyIPropertyStorage|Revert|Discards any changes that have been made
PyObject *PyIPropertyStorage::Revert(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    if (!PyArg_ParseTuple(args, ":Revert"))
        return NULL;
    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->Revert();
    PY_INTERFACE_POSTCALL;

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    Py_INCREF(Py_None);
    return Py_None;
}

// @pymethod <o PyIEnumSTATPROPSTG>|PyIPropertyStorage|Enum|Creates an enumerator for properties in the property set
PyObject *PyIPropertyStorage::Enum(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    IEnumSTATPROPSTG *ppenum;
    if (!PyArg_ParseTuple(args, ":Enum"))
        return NULL;
    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->Enum(&ppenum);
    PY_INTERFACE_POSTCALL;

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    return PyCom_PyObjectFromIUnknown(ppenum, IID_IEnumSTATPROPSTG, FALSE);
}

// @pymethod |PyIPropertyStorage|SetTimes|Sets the creation, last access, and modification time
// @comm Some property sets do not support these times.
PyObject *PyIPropertyStorage::SetTimes(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    // @pyparm <o PyDateTime>|ctime||Creation time, or None for no change
    // @pyparm <o PyDateTime>|atime||Last access time, or None for no change
    // @pyparm <o PyDateTime>|mtime||Modification time, or None for no change
    PyObject *obctime;
    PyObject *obatime;
    PyObject *obmtime;
    FILETIME ctime, *pctime = NULL;
    FILETIME atime, *patime = NULL;
    FILETIME mtime, *pmtime = NULL;
    if (!PyArg_ParseTuple(args, "OOO:SetTimes", &obctime, &obatime, &obmtime))
        return NULL;
    if (obctime != Py_None) {
        if (!PyWinObject_AsFILETIME(obctime, &ctime))
            return NULL;
        pctime = &ctime;
    }
    if (obatime != Py_None) {
        if (!PyWinObject_AsFILETIME(obatime, &atime))
            return NULL;
        patime = &atime;
    }
    if (obmtime != Py_None) {
        if (!PyWinObject_AsFILETIME(obmtime, &mtime))
            return NULL;
        pmtime = &mtime;
    }

    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->SetTimes(pctime, patime, pmtime);
    PY_INTERFACE_POSTCALL;

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    Py_INCREF(Py_None);
    return Py_None;
}

// @pymethod |PyIPropertyStorage|SetClass|Sets the GUID for the property set
PyObject *PyIPropertyStorage::SetClass(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    // @pyparm <o PyIID>|clsid||Description for clsid
    PyObject *obclsid;
    IID clsid;
    if (!PyArg_ParseTuple(args, "O:SetClass", &obclsid))
        return NULL;
    if (!PyWinObject_AsIID(obclsid, &clsid))
        return NULL;
    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->SetClass(clsid);
    PY_INTERFACE_POSTCALL;

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    Py_INCREF(Py_None);
    return Py_None;
}

// @pymethod tuple|PyIPropertyStorage|Stat|Returns various infomation about the property set
// @rdesc Returns a tuple representing a STATPROPSETSTG struct.
PyObject *PyIPropertyStorage::Stat(PyObject *self, PyObject *args)
{
    IPropertyStorage *pIPS = GetI(self);
    if (pIPS == NULL)
        return NULL;
    if (!PyArg_ParseTuple(args, ":Stat"))
        return NULL;
    STATPROPSETSTG p;
    HRESULT hr;
    PY_INTERFACE_PRECALL;
    hr = pIPS->Stat(&p);
    PY_INTERFACE_POSTCALL;

    if (FAILED(hr))
        return PyCom_BuildPyException(hr, pIPS, IID_IPropertyStorage);
    return PyCom_PyObjectFromSTATPROPSETSTG(&p);
}

// @object PyIPropertyStorage|Structured storage object that contains a set of properties.
//	Supports iteration to list properties.
static struct PyMethodDef PyIPropertyStorage_methods[] = {
    {"ReadMultiple", PyIPropertyStorage::ReadMultiple,
     1},  // @pymeth ReadMultiple|Reads specified properties from the current property set.
    {"WriteMultiple", PyIPropertyStorage::WriteMultiple,
     1},  // @pymeth WriteMultiple|Creates or modifies properties in the property set
    {"DeleteMultiple", PyIPropertyStorage::DeleteMultiple,
     1},  // @pymeth DeleteMultiple|Deletes properties from the property set
    {"ReadPropertyNames", PyIPropertyStorage::ReadPropertyNames,
     1},  // @pymeth ReadPropertyNames|Retrieves any existing string names for the specified property identifiers.
    {"WritePropertyNames", PyIPropertyStorage::WritePropertyNames,
     1},  // @pymeth WritePropertyNames|Assigns string names to a specified array of property IDs in the current
          // property set.
    {"DeletePropertyNames", PyIPropertyStorage::DeletePropertyNames,
     1},  // @pymeth DeletePropertyNames|Removes property names from specified properties.
    {"Commit", PyIPropertyStorage::Commit, 1},  // @pymeth Commit|Persists the property set to its base storage
    {"Revert", PyIPropertyStorage::Revert, 1},  // @pymeth Revert|Discards any changes that have been made
    {"Enum", PyIPropertyStorage::Enum, 1},      // @pymeth Enum|Creates an enumerator for properties in the property set
    {"SetTimes", PyIPropertyStorage::SetTimes,
     1},  // @pymeth SetTimes|Sets the creation, last access, and modification time
    {"SetClass", PyIPropertyStorage::SetClass, 1},  // @pymeth SetClass|Sets the GUID for the property set
    {"Stat", PyIPropertyStorage::Stat, 1},          // @pymeth Stat|Returns various infomation about the property set
    {NULL}};

PyComEnumProviderTypeObject PyIPropertyStorage::type("PyIPropertyStorage", &PyIUnknown::type,
                                                     sizeof(PyIPropertyStorage), PyIPropertyStorage_methods,
                                                     GET_PYCOM_CTOR(PyIPropertyStorage), "Enum");

// ---------------------------------------------------
//
// Gateway Implementation
STDMETHODIMP PyGPropertyStorage::ReadMultiple(
    /* [in] */ ULONG cpspec,
    /* [size_is][in] */ const PROPSPEC rgpspec[],
    /* [size_is][out] */ PROPVARIANT rgpropvar[])
{
    PY_GATEWAY_METHOD;
    PyObject *obpropspecs = PyWinObject_FromPROPSPECs(rgpspec, cpspec);
    if (obpropspecs == NULL)
        return MAKE_PYCOM_GATEWAY_FAILURE_CODE("ReadMultiple");
    PyObject *result;
    HRESULT hr = InvokeViaPolicy("ReadMultiple", &result, "(O)", obpropspecs);
    Py_DECREF(obpropspecs);
    if (FAILED(hr))
        return hr;
    if (!PyObject_AsPreallocatedPROPVARIANTs(result, rgpropvar, cpspec))
        hr = MAKE_PYCOM_GATEWAY_FAILURE_CODE("ReadMultiple");
    Py_DECREF(result);
    return hr;
}

STDMETHODIMP PyGPropertyStorage::WriteMultiple(
    /* [in] */ ULONG cpspec,
    /* [size_is][in] */ const PROPSPEC rgpspec[],
    /* [size_is][in] */ const PROPVARIANT rgpropvar[],
    /* [in] */ PROPID propidNameFirst)
{
    PY_GATEWAY_METHOD;
    HRESULT hr;
    {
        TmpPyObject obpropspecs = PyWinObject_FromPROPSPECs(rgpspec, cpspec);
        if (obpropspecs == NULL)
            return MAKE_PYCOM_GATEWAY_FAILURE_CODE("WriteMultiple");
        TmpPyObject obpropvars = PyObject_FromPROPVARIANTs((PROPVARIANT *)rgpropvar, cpspec);
        if (obpropvars == NULL)
            return MAKE_PYCOM_GATEWAY_FAILURE_CODE("WriteMultiple");
        hr = InvokeViaPolicy("WriteMultiple", NULL, "OOk", obpropspecs, obpropvars, propidNameFirst);
    }
    return hr;
}

STDMETHODIMP PyGPropertyStorage::DeleteMultiple(
    /* [in] */ ULONG cpspec,
    /* [size_is][in] */ const PROPSPEC rgpspec[])
{
    PY_GATEWAY_METHOD;
    PyObject *obpropspecs = PyWinObject_FromPROPSPECs(rgpspec, cpspec);
    if (obpropspecs == NULL)
        return MAKE_PYCOM_GATEWAY_FAILURE_CODE("DeleteMultiple");
    HRESULT hr = InvokeViaPolicy("DeleteMultiple", NULL, "O", cpspec);
    Py_DECREF(obpropspecs);
    return hr;
}

// Converts a sequence of strings into a caller-allocated array of WCHAR pointers,
// each of which is allocated using CoTaskMemAlloc
BOOL PyWinObject_AsTaskAllocatedWCHARArray(PyObject *str_seq, LPWSTR *wchars, ULONG str_cnt)
{
    TmpPyObject str_tuple;
    ULONG seq_size, tuple_index;
    ZeroMemory(wchars, str_cnt * sizeof(WCHAR *));
    if ((str_tuple = PyWinSequence_Tuple(str_seq, &seq_size)) == NULL)
        return FALSE;
    if (seq_size != str_cnt) {
        PyErr_SetString(PyExc_ValueError, "Sequence not of required length");
        return FALSE;
    }
    for (tuple_index = 0; tuple_index < str_cnt; tuple_index++) {
        PyObject *tuple_item = PyTuple_GET_ITEM((PyObject *)str_tuple, tuple_index);
        if (!PyWinObject_AsTaskAllocatedWCHAR(tuple_item, &wchars[tuple_index], FALSE)) {
            for (tuple_index = 0; tuple_index < str_cnt; tuple_index++)
                if (wchars[tuple_index])
                    CoTaskMemFree(wchars[tuple_index]);
            return FALSE;
        }
    }
    return TRUE;
}

STDMETHODIMP PyGPropertyStorage::ReadPropertyNames(
    /* [in] */ ULONG cpropid,
    /* [size_is][in] */ const PROPID rgpropid[],
    /* [size_is][out] */ LPOLESTR rglpwstrName[])
{
    PY_GATEWAY_METHOD;
    PyObject *obpropids = PyWinObject_FromPROPIDs(rgpropid, cpropid);
    if (obpropids == NULL)
        return MAKE_PYCOM_GATEWAY_FAILURE_CODE("ReadPropertyNames");
    PyObject *result;
    HRESULT hr = InvokeViaPolicy("ReadPropertyNames", &result, "O", obpropids);
    Py_DECREF(obpropids);
    if (FAILED(hr))
        return hr;
    // Process the Python results, and convert back to the real params
    if (!PyWinObject_AsTaskAllocatedWCHARArray(result, rglpwstrName, cpropid))
        hr = MAKE_PYCOM_GATEWAY_FAILURE_CODE("ReadPropertyNames");
    Py_DECREF(result);
    return hr;
}

PyObject *PyWinObject_FromWCHARArray(const LPOLESTR names[], ULONG cnames)
{
    PyObject *ret = PyTuple_New(cnames);
    if (ret == NULL)
        return NULL;
    PyObject *item;
    for (ULONG i = 0; i < cnames; i++) {
        item = PyWinObject_FromWCHAR(names[i]);
        if (item == NULL) {
            Py_DECREF(ret);
            return NULL;
        }
        PyTuple_SET_ITEM(ret, i, item);
    }
    return ret;
}

STDMETHODIMP PyGPropertyStorage::WritePropertyNames(
    /* [in] */ ULONG cpropid,
    /* [size_is][in] */ const PROPID rgpropid[],
    /* [size_is][in] */ const LPOLESTR rglpwstrName[])
{
    PY_GATEWAY_METHOD;
    {  // Scope so Tmp objects are DECREF'ed before lock released
        TmpPyObject obpropids = PyWinObject_FromPROPIDs(rgpropid, cpropid);
        if (obpropids == NULL)
            return MAKE_PYCOM_GATEWAY_FAILURE_CODE("WritePropertyNames");
        TmpPyObject obnames = PyWinObject_FromWCHARArray(rglpwstrName, cpropid);
        if (obnames == NULL)
            return MAKE_PYCOM_GATEWAY_FAILURE_CODE("WritePropertyNames");
        return InvokeViaPolicy("WritePropertyNames", NULL, "OO", obpropids, obnames);
    }
}

STDMETHODIMP PyGPropertyStorage::DeletePropertyNames(
    /* [in] */ ULONG cpropid,
    /* [size_is][in] */ const PROPID rgpropid[])
{
    PY_GATEWAY_METHOD;
    PyObject *obpropids = PyWinObject_FromPROPIDs(rgpropid, cpropid);
    if (obpropids == NULL)
        return MAKE_PYCOM_GATEWAY_FAILURE_CODE("DeletePropertyNames");
    HRESULT hr = InvokeViaPolicy("DeletePropertyNames", NULL, "O", obpropids);
    Py_DECREF(obpropids);
    return hr;
}

STDMETHODIMP PyGPropertyStorage::Commit(
    /* [in] */ DWORD grfCommitFlags)
{
    PY_GATEWAY_METHOD;
    HRESULT hr = InvokeViaPolicy("Commit", NULL, "l", grfCommitFlags);
    return hr;
}

STDMETHODIMP PyGPropertyStorage::Revert(void)
{
    PY_GATEWAY_METHOD;
    HRESULT hr = InvokeViaPolicy("Revert", NULL);
    return hr;
}

STDMETHODIMP PyGPropertyStorage::Enum(
    /* [out] */ IEnumSTATPROPSTG **ppenum)
{
    PY_GATEWAY_METHOD;
    if (ppenum == NULL)
        return E_POINTER;
    PyObject *result;
    HRESULT hr = InvokeViaPolicy("Enum", &result);
    if (FAILED(hr))
        return hr;
    // Process the Python results, and convert back to the real params
    if (!PyCom_InterfaceFromPyInstanceOrObject(result, IID_IEnumSTATPROPSTG, (void **)ppenum, FALSE))
        hr = MAKE_PYCOM_GATEWAY_FAILURE_CODE("Enum");
    Py_DECREF(result);
    return hr;
}

STDMETHODIMP PyGPropertyStorage::SetTimes(
    /* [in] */ const FILETIME *pctime,
    /* [in] */ const FILETIME *patime,
    /* [in] */ const FILETIME *pmtime)
{
    PY_GATEWAY_METHOD;
    HRESULT hr;
    {
        TmpPyObject obctime, obatime, obmtime;
        if (pctime) {
            obctime = PyWinObject_FromFILETIME(*pctime);
            if (obctime == NULL)
                return MAKE_PYCOM_GATEWAY_FAILURE_CODE("SetTimes");
        }
        else {
            Py_INCREF(Py_None);
            obctime = Py_None;
        }
        if (patime) {
            obatime = PyWinObject_FromFILETIME(*patime);
            if (obatime == NULL)
                return MAKE_PYCOM_GATEWAY_FAILURE_CODE("SetTimes");
        }
        else {
            Py_INCREF(Py_None);
            obatime = Py_None;
        }
        if (pmtime) {
            obmtime = PyWinObject_FromFILETIME(*pmtime);
            if (obmtime == NULL)
                return MAKE_PYCOM_GATEWAY_FAILURE_CODE("SetTimes");
        }
        else {
            Py_INCREF(Py_None);
            obmtime = Py_None;
        }
        hr = InvokeViaPolicy("SetTimes", NULL, "OOO", obctime, obatime, obmtime);
    }
    return hr;
}

STDMETHODIMP PyGPropertyStorage::SetClass(
    /* [in] */ REFCLSID clsid)
{
    PY_GATEWAY_METHOD;
    PyObject *obclsid;
    obclsid = PyWinObject_FromIID(clsid);
    HRESULT hr = InvokeViaPolicy("SetClass", NULL, "O", obclsid);
    Py_XDECREF(obclsid);
    return hr;
}

STDMETHODIMP PyGPropertyStorage::Stat(
    /* [out] */ STATPROPSETSTG *pstatpsstg)
{
    PY_GATEWAY_METHOD;
    PyObject *result;
    HRESULT hr = InvokeViaPolicy("Stat", &result);
    if (FAILED(hr))
        return hr;
    if (!PyCom_PyObjectAsSTATPROPSETSTG(result, pstatpsstg))
        hr = MAKE_PYCOM_GATEWAY_FAILURE_CODE("Stat");

    Py_DECREF(result);
    return hr;
}

#endif  // NO_PYCOM_IPROPERTYSTORAGE
