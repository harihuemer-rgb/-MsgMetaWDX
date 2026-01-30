
// MsgMetaWDX.cpp - WDX-Plugin für MSG-Metadaten (x64, Unicode)
// Build: Windows DLL; Link: MAPI32.lib, Ole32.lib, Uuid.lib
// Felder: From, To, Subject, Sent, Received, Message-ID, HasAttachments
// Hinweis: Outlook/Extended MAPI muss installiert sein.

#include <windows.h>
#include <string>
#include <mapi.h>
#include <mapiutil.h>
#include <mapidefs.h>
#include <mapitags.h>
#include <imessage.h>
#include <objbase.h>
#include <objidl.h>

#pragma comment(lib, "MAPI32.lib")
#pragma comment(lib, "Ole32.lib")

// --- WDX Feldtypen (Minimaldefinition für Build) ---
#ifndef FT_DEFS
#define FT_DEFS
#define ft_nomorefields   0
#define ft_numeric_32     1
#define ft_numeric_64     2
#define ft_numeric_floating 3
#define ft_date           4
#define ft_time           5
#define ft_boolean        6
#define ft_multiplechoice 7
#define ft_string         8
#define ft_fulltext       9
#define ft_datetime       10
#define ft_stringw        11
#define ft_fulltextw      12

#define ft_nosuchfield   -1
#define ft_fileerror     -2
#define ft_fieldempty    -3
#endif

// ---- Felder ----
enum FieldId {
    F_FROM = 0, F_TO, F_SUBJECT, F_SENT, F_RECEIVED, F_MSGID, F_HASATTACH, F_COUNT
};

struct MsgMeta {
    std::wstring From, To, Subject, MessageId;
    FILETIME Sent{}, Received{};
    BOOL HasAttach = FALSE;
};

static bool g_comInited = false;
static bool g_mapiInited = false;

static HRESULT EnsureCOM() {
    if (!g_comInited) {
        HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        if (SUCCEEDED(hr) || hr == S_FALSE) { g_comInited = true; return S_OK; }
        return hr;
    }
    return S_OK;
}
static HRESULT EnsureMAPI() {
    if (!g_mapiInited) {
        HRESULT hr = MAPIInitialize(nullptr);
        if (SUCCEEDED(hr)) g_mapiInited = true;
        return hr;
    }
    return S_OK;
}
static std::wstring A2W(const char* s) {
    if (!s) return L"";
    int n = MultiByteToWideChar(CP_ACP, 0, s, -1, nullptr, 0);
    std::wstring w; w.resize(n ? n - 1 : 0);
    if (n > 1) MultiByteToWideChar(CP_ACP, 0, s, -1, &w[0], n);
    return w;
}
static std::wstring PropToWString(const SPropValue* p) {
    if (!p) return L"";
    auto t = PROP_TYPE(p->ulPropTag);
    if (t == PT_UNICODE)  return p->Value.lpszW ? p->Value.lpszW : L"";
    if (t == PT_STRING8)  return A2W(p->Value.lpszA);
    return L"";
}
static HRESULT GetOnePropW(IMAPIProp* p, ULONG tagW, ULONG tagA, std::wstring& out) {
    SPropValue* v = nullptr;
    HRESULT hr = HrGetOneProp(p, tagW, &v);
    if (FAILED(hr)) hr = HrGetOneProp(p, tagA, &v);
    if (SUCCEEDED(hr) && v) { out = PropToWString(v); MAPIFreeBuffer(v); return S_OK; }
    if (v) MAPIFreeBuffer(v);
    return hr;
}

static HRESULT ReadMsgMeta(const wchar_t* file, MsgMeta& m) {
    HRESULT hr = EnsureCOM();   if (FAILED(hr)) return hr;
    hr = EnsureMAPI();          if (FAILED(hr)) return hr;

    CComPtr<IStorage> stg;
    hr = StgOpenStorage(file, nullptr, STGM_READ | STGM_SHARE_DENY_WRITE, nullptr, 0, &stg);
    if (FAILED(hr)) return hr;

    CComPtr<IMalloc> mallocIF;
    hr = CoGetMalloc(1, &mallocIF);
    if (FAILED(hr)) return hr;

    LPMSGSESS sess = nullptr;
    hr = OpenIMsgSession(mallocIF, 0, &sess);
    if (FAILED(hr)) return hr;

    IMessage* msg = nullptr;
    hr = OpenIMsgOnIStg(
        sess, MAPIAllocateBuffer, MAPIAllocateMore, MAPIFreeBuffer,
        mallocIF, nullptr, stg, nullptr, 0, 0, &msg
    );
    if (FAILED(hr)) { CloseIMsgSession(sess); return hr; }

    std::wstring senderName, senderEmail;
    GetOnePropW(msg, PR_SENDER_NAME_W,            PR_SENDER_NAME_A,            senderName);
    GetOnePropW(msg, PR_SENDER_EMAIL_ADDRESS_W,   PR_SENDER_EMAIL_ADDRESS_A,   senderEmail);
    m.From = senderEmail.empty() ? senderName :
             (senderName.empty() ? senderEmail : (senderName + L" <" + senderEmail + L">");

    GetOnePropW(msg, PR_DISPLAY_TO_W,             PR_DISPLAY_TO_A,             m.To);
    GetOnePropW(msg, PR_SUBJECT_W,                PR_SUBJECT_A,                m.Subject);
    GetOnePropW(msg, PR_INTERNET_MESSAGE_ID_W,    PR_INTERNET_MESSAGE_ID_A,    m.MessageId);

    SPropValue* v = nullptr;
    if (SUCCEEDED(HrGetOneProp(msg, PR_CLIENT_SUBMIT_TIME, &v))) {
        m.Sent = v->Value.ft; MAPIFreeBuffer(v);
    } else ZeroMemory(&m.Sent, sizeof(FILETIME));

    if (SUCCEEDED(HrGetOneProp(msg, PR_MESSAGE_DELIVERY_TIME, &v))) {
        m.Received = v->Value.ft; MAPIFreeBuffer(v);
    } else ZeroMemory(&m.Received, sizeof(FILETIME));

    if (SUCCEEDED(HrGetOneProp(msg, PR_HASATTACH, &v))) {
        m.HasAttach = v->Value.b; MAPIFreeBuffer(v);
    } else m.HasAttach = FALSE;

    msg->Release();
    CloseIMsgSession(sess);
    return S_OK;
}

extern "C" {

int __stdcall ContentGetDetectString(char* DetectString, int maxlen) {
    const char* s = "EXT=\"MSG\"";
    lstrcpynA(DetectString, s, maxlen);
    return 0;
}

int __stdcall ContentGetSupportedField(int FieldIndex, char* FieldName, char* Units, int maxlen) {
    struct { const char* name; int type; } F[] = {
        { "From",       ft_stringw },
        { "To",         ft_stringw },
        { "Subject",    ft_stringw },
        { "Sent",       ft_datetime },
        { "Received",   ft_datetime },
        { "Message-ID", ft_stringw },
        { "HasAttachments", ft_boolean },
    };
    if (FieldIndex < 0 || FieldIndex >= (int)F_COUNT) return ft_nomorefields;
    lstrcpynA(FieldName, F[FieldIndex].name, maxlen);
    if (Units) *Units = 0;
    return F[FieldIndex].type;
}

static int FillValueW(const wchar_t* FileName, int FieldIndex, void* FieldValue, int maxlen) {
    MsgMeta m;
    if (FAILED(ReadMsgMeta(FileName, m))) return ft_fileerror;

    switch (FieldIndex) {
    case F_FROM:       lstrcpynW((wchar_t*)FieldValue, m.From.c_str(),  maxlen/2); return ft_stringw;
    case F_TO:         lstrcpynW((wchar_t*)FieldValue, m.To.c_str(),    maxlen/2); return ft_stringw;
    case F_SUBJECT:    lstrcpynW((wchar_t*)FieldValue, m.Subject.c_str(), maxlen/2); return ft_stringw;
    case F_MSGID:      lstrcpynW((wchar_t*)FieldValue, m.MessageId.c_str(), maxlen/2);
                       return m.MessageId.empty() ? ft_fieldempty : ft_stringw;
    case F_SENT:       *(FILETIME*)FieldValue = m.Sent;     return (m.Sent.dwHighDateTime||m.Sent.dwLowDateTime)? ft_datetime : ft_fieldempty;
    case F_RECEIVED:   *(FILETIME*)FieldValue = m.Received; return (m.Received.dwHighDateTime||m.Received.dwLowDateTime)? ft_datetime : ft_fieldempty;
    case F_HASATTACH:  *(int*)FieldValue = m.HasAttach ? 1 : 0; return ft_boolean;
    default:           return ft_nosuchfield;
    }
}

int __stdcall ContentGetValueW(WCHAR* FileName, int FieldIndex, int /*UnitIndex*/,
                               void* FieldValue, int maxlen, int /*flags*/) {
    return FillValueW(FileName, FieldIndex, FieldValue, maxlen);
}

int __stdcall ContentGetValue(char* FileName, int FieldIndex, int UnitIndex,
                              void* FieldValue, int maxlen, int flags) {
    int wlen = MultiByteToWideChar(CP_ACP, 0, FileName, -1, nullptr, 0);
    std::wstring w; w.resize(wlen ? wlen - 1 : 0);
    if (wlen > 1) MultiByteToWideChar(CP_ACP, 0, FileName, -1, &w[0], wlen);
    return ContentGetValueW((WCHAR*)w.c_str(), FieldIndex, UnitIndex, FieldValue, maxlen, flags);
}

void __stdcall ContentPluginUnloading(void) {
    if (g_mapiInited) { MAPIUninitialize(); g_mapiInited = false; }
    if (g_comInited)  { CoUninitialize();    g_comInited  = false; }
}

}
