# ListViewSubItemControls v1.1
Undocumented ListView SubItem Controls Demo

![ScreenShot](https://github.com/user-attachments/assets/f3cf5881-e693-4367-a99e-0a3c702e30b6)

This project has been my white whale. Back in 2015 I started a series of articles on undocumented ListView features available in Windows Vista+: [Footer Items](https://www.vbforums.com/showthread.php?798159-VB6-Vista-Undocumented-ListView-feature-Footer-items), [Subsetted Groups](https://www.vbforums.com/showthread.php?798321-VB6-Vista-Undocumented-ListView-feature-Subsetted-Groups-(simple-no-TLB)), [Groups in Virtual Mode](https://www.vbforums.com/showthread.php?808981-VB6-Vista-Undocumented-ListView-feature-Groups-in-Virtual-Mode), [Column Backcolors](https://www.vbforums.com/showthread.php?869049-VB6-Undocumented-ListView-feature-Highlight-column), and [Explorer-style selection](https://www.vbforums.com/showthread.php?894841-VB6-Win7-Undocumented-ListView-Feature-Multiselect-in-first-column-like-Explorer). But the coolest undocumented feature of all was the automatic subitem controls shown in the picture above. I just could not get it. The work was based on a fully working [project by Timo Kunze](https://www.codeproject.com/Articles/35197/Undocumented-List-View-Features), but even though I had this C++ sample that worked, every effort to port it to VB6 failed. Weeks were spent on it. Then at least half a dozen major efforts over the following decade of a few days. Not willing to give it up, I tried again starting 2 days ago, only this time instead of trying to fix the giant mess of spaghetti code packed with debugging stuff and the remnants of numerous different approaches, I started over completely from scratch and tried to make the port as line-by-line identical as possible...

🥳🥳 ** IT WORKED ** 🥳🥳

I'll no doubt be digging into the old code to find out exactly what I could have possibly missed in all the other failed attempts, which only ever got as far as glitched rendering of one or two controls followed by a hard crash. But the bottom line is now every control is working perfectly! In both 32 and 64bit! In future versions I'll explore the control types not used in Timo's demo.

**Requirements**
- Windows 7+ (Vista+ could be supported by switching the IListView version, but it's not done here in v1.0).
- Windows Development Library for twinBASIC v9.1+
- Common Controls 6.0 enabled by manifest

**Updates**
v1.1 (22 Jun 2025)
Now supports toggling to Tiles view to show how they work great there too:
![image](https://github.com/user-attachments/assets/4610bd77-b391-4e76-a4e6-dc4342f587b3)


**How it works**
This technique is based around the undocumented `ISubItemCallback` interface:

```vba
[InterfaceId("11A66240-5489-42C2-AEBF-286FC831524C")]
[OleAutomation(False)]
Interface ISubItemCallback Extends stdole.IUnknown
    Sub GetSubItemTitle(ByVal subitemIndex As Long, ByVal lpszBuffer As LongPtr, ByVal BufferSize As Long)
    Sub GetSubItemControl(ByVal itemIndex As Long, ByVal subItemIndex As Long, requiredInterface As UUID, ppObject As Any)
    Sub BeginSubItemEdit(ByVal itemIndex As Long, ByVal subItemIndex As Long, ByVal mode As Long, requiredInterface As UUID, ppObject As Any)
    Sub EndSubItemEdit(ByVal itemIndex As Long, ByVal subItemIndex As Long, ByVal mode As Long, ByVal ppc As IPropertyControl)
    Sub BeginGroupEdit(ByVal groupIndex As Long, requiredInterface As UUID, ppObject As Any)
    Sub EndGroupEdit(ByVal groupIndex As Long, ByVal mode As Long, ByVal pPropertyControl As IPropertyControl)
    Sub OnInvokeVerb(ByVal itemIndex As Long, ByVal pVerb As LongPtr)
End Interface
```

We implement in our Form then set it as the callback object via `IListView`'s `SetSubItemCallback` method. The initial values of the controls we just set by the subitem text of the listview items, e.g. 46 for 46% on the percent bar, or the numeric value of a `FILETIME` for the Date/Time control. The control in the picture that says 'Center weighted average' is actually a EXIF property the shell displays for photos... it automatically populates the values just by giving it an `IPropertyDescription` for System.Photo.MeteringMode, and we just provide an index.

The `BeginSubItemEdit` and identically handled `GetSubItemControl` callback methods are where the interesting part happens. This is a difficult interface to use with `Implements` because these take `void**` (`As Any`) arguments. twinBASIC lets us implement these methods as-is by using `ByRef LongPtr`; the key is we use `CoCreateInstance` to create an object for these controls on the argument, only when the subitemIndex is 1.

```vba
        Select Case itemIndex
            Case 0
                hr = CoCreateInstance(CLSID_CInPlaceMLEditBoxControl, Nothing, CLSCTX_INPROC_SERVER, requiredInterface, ppObject)
                If pPropertyDescription Is Nothing Then
                    PSGetPropertyDescriptionByName("System.Generic.String", IID_IPropertyDescription, pPropertyDescription)
                End If
            Case 1
                hr = CoCreateInstance(CLSID_CCustomDrawPercentFullControl, Nothing, CLSCTX_INPROC_SERVER, requiredInterface, ppObject)
            Case 2
                hr = CoCreateInstance(CLSID_CRatingControl, Nothing, CLSCTX_INPROC_SERVER, requiredInterface, ppObject)
            Case 3
                If IsEqualIID(requiredInterface, IID_IDrawPropertyControl) Then
                    hr = CoCreateInstance(CLSID_CStaticPropertyControl, Nothing, CLSCTX_INPROC_SERVER, requiredInterface, ppObject)
                Else
                    hr = CoCreateInstance(CLSID_CInPlaceEditBoxControl, Nothing, CLSCTX_INPROC_SERVER, requiredInterface, ppObject)
                End If
                If pPropertyDescription Is Nothing Then
                    PSGetPropertyDescriptionByName("System.Generic.String", IID_IPropertyDescription, pPropertyDescription)
                End If

```

etc.

After that there's some voodoo regarding visual styles I don't fully understand; I think it's just getting the strings for the window themes from the `GetProp` accessible locations they're stored in; but Timo doesn't explain why we don't just pass "explorer" like the main ListView control. Finally, and this is where I think earlier efforts went wrong, is the method for storing and editing values with `IPropertyValue`. I didn't want tB doing anything behind my back, so almost everything there now uses a manually defined UDT version of `PROPVARIANT`. We have a class that implements it, and we populate it with the text values of the subitem coerced into their proper `PROPVARIANT` data type; `VT_LPWSTR`   types are an epic pain and this time around I think battle-tested helpers I had for it made a difference.

```vba
            Dim pBuffer As LongPtr = HeapAlloc(GetProcessHeap(), 0, (1024 + 1) * 2 /* sizeof(WCHAR) */)
            If pBuffer Then
                Dim item As LVITEMW
                item.iSubItem = subItemIndex
                item.cchTextMax = 1024
                item.pszText = pBuffer
                SendMessage(hLV, LVM_GETITEMTEXTW, itemIndex, item)
                If (itemIndex = 1) Or (itemIndex = 2) Or (itemIndex = 7) Then
                    Dim tmp As Variant
                    PropVariantInit(tmp)
                    InitPropVariantFromString(item.pszText, tmp)
                    PropVariantChangeType(ByVal pPropertyValue, tmp, 0, VT_UI4)
                    PropVariantClear(tmp)
...
            Dim pPropertyValueObj As IPropertyValue
            Set pPropertyValueObj = New IPropertyValueImpl
            pPropertyValueObj.InitValue(propertyValue)
            If pPropertyDescription IsNot Nothing Then
                pControl.Initialize(pPropertyDescription, 0)
            End If
            pControl.SetValue(pPropertyValueObj)
            pControl.SetTextColor(textColor)
            If hFont Then
                pControl.SetFont(hFont)
            End If
```

etc.

When the values are changed through the control, it sends the `EndSubItemEdit` and we take the `IPropertyValue` interface and turn it back into a String we store as the item text:

```vba
    Private Sub ISubItemCallback_EndSubItemEdit(ByVal itemIndex As Long, ByVal subItemIndex As Long, ByVal mode As Long, ByVal ppc As IPropertyControl) Implements ISubItemCallback.EndSubItemEdit
...
        Dim modified As BOOL
        ppc.IsModified(modified)
        If modified Then
            Dim pPropertyValue As IPropertyValue
            ppc.GetValue(IID_IPropertyValue, pPropertyValue)
            If SUCCEEDED(Err.LastHresult) Then
                Dim propertyValue As PROPVARIANT
                PropVariantInit(propertyValue)
                pPropertyValue.GetValue(propertyValue)
                If SUCCEEDED(Err.LastHresult) Then
                    Dim pBuffer As LongPtr
                    If SUCCEEDED(PropVariantToStringAlloc(propertyValue, pBuffer)) AndAlso (pBuffer <> 0) Then
                        LVSetItemText(itemIndex, subItemIndex, pBuffer)
                        CoTaskMemFree(pBuffer)
                    End If
                    PropVariantClear(propertyValue)
                End If
            End If
        End If
```

So that's the broad strokes. There's obviously a ton of details I've left out, so grab the code and dig in! Later this summer I hope to bring these controls to ucShellBrowse, as they're far more stable than my current method of creating a bunch of new windows and drawing the stars myself in `WM_PAINT` handlers. 😄
