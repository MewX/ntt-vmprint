/*++

Copyright (c) 2005 Microsoft Corporation

All rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.

File Name:

   coldmptcnv.h

Abstract:

   PageSourceColorProfile devmode <-> PrintTicket conversion class definition.
   The class defines a common data representation between the DevMode (GPD) and PrintTicket
   representations and implements the conversion and validation methods required
   by CFeatureDMPTConvert.

--*/

#pragma once

#include "ftrdmptcnv.h"
#include "cmprofpthndlr.h"
#include "cmprofiledata.h"

class CColorProfileDMPTConv :
    public CFeatureDMPTConvert<XDPrintSchema::PageSourceColorProfile::PageSourceColorProfileData>
{
public:
    CColorProfileDMPTConv();

    virtual ~CColorProfileDMPTConv();

private:
    HRESULT
    GetPTDataSettingsFromDM(
        _In_    PDEVMODE         pDevmode,
        _In_    ULONG            cbDevmode,
        _In_    PVOID            pPrivateDevmode,
        _In_    ULONG            cbDrvPrivateSize,
        _Out_   XDPrintSchema::PageSourceColorProfile::PageSourceColorProfileData* pDataSettings
        );

    HRESULT
    MergePTDataSettingsWithPT(
        _In_    IXMLDOMDocument2* pPrintTicket,
        _Inout_ XDPrintSchema::PageSourceColorProfile::PageSourceColorProfileData*  pDrvSettings
        );

    HRESULT
    SetPTDataInDM(
        _In_    CONST XDPrintSchema::PageSourceColorProfile::PageSourceColorProfileData& drvSettings,
        _Inout_ PDEVMODE pDevmode,
        _In_    ULONG    cbDevmode,
        _Inout_ PVOID    pPrivateDevmode,
        _In_    ULONG    cbDrvPrivateSize
        );

    HRESULT
    SetPTDataInPT(
        _In_    CONST XDPrintSchema::PageSourceColorProfile::PageSourceColorProfileData& drvSettings,
        _Inout_ IXMLDOMDocument2* pPrintTicket
        );

    HRESULT STDMETHODCALLTYPE
    CompletePrintCapabilities(
        _In_opt_ IXMLDOMDocument2*,
        _Inout_  IXMLDOMDocument2* pPrintCapabilities
        );
};

