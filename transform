#ifndef UNICODE
#define UNICODE
#endif

#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif

#include <windows.h>
#include <commdlg.h>
#include <ocidl.h>
#include <objidl.h>
#include <oleauto.h>
#include <propidl.h>
#include <shellapi.h>
#include <wincodec.h>

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cwctype>
#include <limits>
#include <memory>
#include <sstream>
#include <string>
#include <vector>

enum ControlId {
    IDC_INPUT_EDIT = 1001,
    IDC_INPUT_BUTTON,
    IDC_OUTPUT_EDIT,
    IDC_OUTPUT_BUTTON,
    IDC_TARGET_EDIT,
    IDC_COMPRESS_BUTTON,
    IDC_HINT_STATIC,
    IDC_INFO_EDIT
};

constexpr UINT WM_COMPRESS_DONE = WM_APP + 1;
constexpr int kPreferredQuality = 82;
constexpr int kHardMinQuality = 12;

enum class AppLanguage {
    Chinese,
    English
};

AppLanguage g_appLanguage = AppLanguage::English;

AppLanguage DetectAppLanguage() {
    LANGID language = GetThreadUILanguage();
    if (language == 0) {
        language = GetUserDefaultUILanguage();
    }

    return PRIMARYLANGID(language) == LANG_CHINESE ? AppLanguage::Chinese
                                                   : AppLanguage::English;
}

const wchar_t* Tr(const wchar_t* zh, const wchar_t* en) {
    return g_appLanguage == AppLanguage::Chinese ? zh : en;
}

const wchar_t* AppTitle() {
    return Tr(L"\u8F7B\u91CF\u56FE\u7247\u538B\u7F29\u5DE5\u5177", L"Light Image Compressor");
}

const wchar_t* NoticeCaption() {
    return Tr(L"\u63D0\u793A", L"Notice");
}

const wchar_t* CompressionFailedCaption() {
    return Tr(L"\u538B\u7F29\u5931\u8D25", L"Compression Failed");
}

const wchar_t* InputFileFilter() {
    return Tr(
        L"\u56FE\u7247\u6587\u4EF6\0*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff\0"
        L"\u6240\u6709\u6587\u4EF6\0*.*\0",
        L"Image Files\0*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff\0All Files\0*.*\0");
}

const wchar_t* OutputFileFilter() {
    return Tr(
        L"JPEG \u56FE\u7247\0*.jpg;*.jpeg\0\u6240\u6709\u6587\u4EF6\0*.*\0",
        L"JPEG Image\0*.jpg;*.jpeg\0All Files\0*.*\0");
}

struct ImageData {
    UINT width = 0;
    UINT height = 0;
    std::vector<BYTE> pixels;
};

struct CompressionSettings {
    std::wstring inputPath;
    std::wstring outputPath;
    unsigned long long targetBytes = 0;
};

struct CompressionResult {
    bool success = false;
    bool targetMet = false;
    std::wstring message;
    unsigned long long inputBytes = 0;
    unsigned long long outputBytes = 0;
    UINT inputWidth = 0;
    UINT inputHeight = 0;
    UINT outputWidth = 0;
    UINT outputHeight = 0;
    int quality = 0;
    double scale = 1.0;
};

struct Candidate {
    bool valid = false;
    std::vector<BYTE> encodedBytes;
    unsigned long long size = 0;
    int quality = 0;
    double scale = 1.0;
    UINT width = 0;
    UINT height = 0;
    double score = 0.0;
};

struct AppState {
    HWND hwnd = nullptr;
    HWND inputLabel = nullptr;
    HWND inputEdit = nullptr;
    HWND inputButton = nullptr;
    HWND outputLabel = nullptr;
    HWND outputEdit = nullptr;
    HWND outputButton = nullptr;
    HWND targetLabel = nullptr;
    HWND targetEdit = nullptr;
    HWND compressButton = nullptr;
    HWND hintStatic = nullptr;
    HWND infoEdit = nullptr;
    HFONT font = nullptr;
    bool busy = false;
    std::wstring lastAutoOutputPath;
};

struct WorkerContext {
    HWND hwnd = nullptr;
    CompressionSettings settings;
};

template <typename T>
void SafeRelease(T*& ptr) {
    if (ptr != nullptr) {
        ptr->Release();
        ptr = nullptr;
    }
}

std::wstring GetWindowString(HWND hwnd) {
    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0) {
        return L"";
    }

    std::wstring text(static_cast<size_t>(length), L'\0');
    GetWindowTextW(hwnd, &text[0], length + 1);
    return text;
}

void SetWindowString(HWND hwnd, const std::wstring& text) {
    SetWindowTextW(hwnd, text.c_str());
}

std::wstring Trimmed(const std::wstring& value) {
    const size_t start = value.find_first_not_of(L" \t\r\n\"");
    if (start == std::wstring::npos) {
        return L"";
    }

    const size_t end = value.find_last_not_of(L" \t\r\n\"");
    return value.substr(start, end - start + 1);
}

std::wstring LowercaseCopy(std::wstring text) {
    std::transform(text.begin(), text.end(), text.begin(),
                   [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return text;
}

bool TryParseDouble(const std::wstring& text, double& value) {
    const std::wstring trimmed = Trimmed(text);
    if (trimmed.empty()) {
        return false;
    }

    wchar_t* end = nullptr;
    const double parsed = std::wcstod(trimmed.c_str(), &end);
    if (end == trimmed.c_str() || !std::isfinite(parsed)) {
        return false;
    }

    while (*end != L'\0') {
        if (!std::iswspace(*end)) {
            return false;
        }
        ++end;
    }

    value = parsed;
    return true;
}

bool FileExists(const std::wstring& path) {
    const DWORD attrs = GetFileAttributesW(path.c_str());
    return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

bool GetFileSizeBytes(const std::wstring& path, unsigned long long& bytes) {
    WIN32_FILE_ATTRIBUTE_DATA data = {};
    if (!GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &data)) {
        return false;
    }

    ULARGE_INTEGER size = {};
    size.LowPart = data.nFileSizeLow;
    size.HighPart = data.nFileSizeHigh;
    bytes = size.QuadPart;
    return true;
}

std::wstring FormatBytes(unsigned long long bytes) {
    std::wostringstream oss;
    oss.setf(std::ios::fixed);
    oss.precision(1);

    if (bytes < 1024ULL) {
        oss.precision(0);
        oss << bytes << L" B";
        return oss.str();
    }

    const double kb = static_cast<double>(bytes) / 1024.0;
    if (kb < 1024.0) {
        oss << kb << L" KB";
        return oss.str();
    }

    const double mb = kb / 1024.0;
    if (mb < 1024.0) {
        oss << mb << L" MB";
        return oss.str();
    }

    const double gb = mb / 1024.0;
    oss << gb << L" GB";
    return oss.str();
}

std::wstring FormatPercent(double value) {
    std::wostringstream oss;
    oss.setf(std::ios::fixed);
    oss.precision(0);
    oss << value * 100.0 << L"%";
    return oss.str();
}

std::wstring ForceJpegExtension(const std::wstring& inputPath) {
    if (inputPath.empty()) {
        return inputPath;
    }

    std::wstring output = inputPath;
    const size_t slashPos = output.find_last_of(L"\\/");
    const size_t dotPos = output.find_last_of(L'.');
    if (dotPos == std::wstring::npos || (slashPos != std::wstring::npos && dotPos < slashPos)) {
        output += L".jpg";
        return output;
    }

    const std::wstring ext = LowercaseCopy(output.substr(dotPos));
    if (ext == L".jpg" || ext == L".jpeg") {
        return output;
    }

    output.erase(dotPos);
    output += L".jpg";
    return output;
}

std::wstring DefaultOutputPath(const std::wstring& inputPath) {
    if (inputPath.empty()) {
        return L"";
    }

    std::wstring output = inputPath;
    const size_t slashPos = output.find_last_of(L"\\/");
    const size_t dotPos = output.find_last_of(L'.');
    if (dotPos == std::wstring::npos || (slashPos != std::wstring::npos && dotPos < slashPos)) {
        return output + L"_compressed.jpg";
    }

    output.erase(dotPos);
    output += L"_compressed.jpg";
    return output;
}

std::wstring FormatSystemError(DWORD errorCode) {
    if (errorCode == 0) {
        return L"";
    }

    LPWSTR buffer = nullptr;
    const DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
                        FORMAT_MESSAGE_IGNORE_INSERTS;
    const DWORD length = FormatMessageW(flags, nullptr, errorCode, 0,
                                        reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
    if (length == 0 || buffer == nullptr) {
        return L"";
    }

    std::wstring message(buffer, length);
    LocalFree(buffer);

    while (!message.empty() &&
           (message.back() == L'\r' || message.back() == L'\n' || message.back() == L'.' ||
            message.back() == L' ')) {
        message.pop_back();
    }

    return message;
}

std::wstring FormatHResult(HRESULT hr) {
    const std::wstring systemMessage = FormatSystemError(static_cast<DWORD>(hr));
    if (!systemMessage.empty()) {
        return systemMessage;
    }

    std::wostringstream oss;
    oss << L"HRESULT 0x" << std::hex << std::uppercase
        << static_cast<unsigned long>(hr & 0xFFFFFFFFUL);
    return oss.str();
}

HRESULT CreateImagingFactory(IWICImagingFactory** factory) {
    if (factory == nullptr) {
        return E_POINTER;
    }

    *factory = nullptr;
    return CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                            IID_IWICImagingFactory, reinterpret_cast<void**>(factory));
}

HRESULT OpenFirstFrame(IWICImagingFactory* factory,
                       const std::wstring& path,
                       IWICBitmapDecoder** decoderOut,
                       IWICBitmapFrameDecode** frameOut) {
    if (factory == nullptr || decoderOut == nullptr || frameOut == nullptr) {
        return E_POINTER;
    }

    *decoderOut = nullptr;
    *frameOut = nullptr;

    IWICBitmapDecoder* decoder = nullptr;
    IWICBitmapFrameDecode* frame = nullptr;

    HRESULT hr = factory->CreateDecoderFromFilename(path.c_str(), nullptr, GENERIC_READ,
                                                    WICDecodeMetadataCacheOnDemand, &decoder);
    if (SUCCEEDED(hr)) {
        hr = decoder->GetFrame(0, &frame);
    }

    if (FAILED(hr)) {
        SafeRelease(frame);
        SafeRelease(decoder);
        return hr;
    }

    *decoderOut = decoder;
    *frameOut = frame;
    return S_OK;
}

USHORT ReadOrientation(IWICBitmapFrameDecode* frame) {
    if (frame == nullptr) {
        return 1;
    }

    USHORT orientation = 1;
    IWICMetadataQueryReader* reader = nullptr;
    if (FAILED(frame->GetMetadataQueryReader(&reader)) || reader == nullptr) {
        return orientation;
    }

    PROPVARIANT value;
    PropVariantInit(&value);
    if (SUCCEEDED(reader->GetMetadataByName(L"/app1/ifd/{ushort=274}", &value))) {
        if (value.vt == VT_UI2) {
            orientation = value.uiVal;
        } else if (value.vt == VT_UI4) {
            orientation = static_cast<USHORT>(value.ulVal);
        }
    }
    PropVariantClear(&value);
    SafeRelease(reader);

    if (orientation < 1 || orientation > 8) {
        orientation = 1;
    }
    return orientation;
}

void ApplyOrientation(ImageData& image, USHORT orientation) {
    if (orientation == 1 || image.width == 0 || image.height == 0 || image.pixels.empty()) {
        return;
    }

    const UINT srcWidth = image.width;
    const UINT srcHeight = image.height;
    const UINT dstWidth = (orientation >= 5 && orientation <= 8) ? srcHeight : srcWidth;
    const UINT dstHeight = (orientation >= 5 && orientation <= 8) ? srcWidth : srcHeight;

    std::vector<BYTE> output(static_cast<size_t>(dstWidth) * static_cast<size_t>(dstHeight) * 4U);

    for (UINT y = 0; y < dstHeight; ++y) {
        for (UINT x = 0; x < dstWidth; ++x) {
            UINT srcX = 0;
            UINT srcY = 0;

            switch (orientation) {
                case 2:
                    srcX = srcWidth - 1U - x;
                    srcY = y;
                    break;
                case 3:
                    srcX = srcWidth - 1U - x;
                    srcY = srcHeight - 1U - y;
                    break;
                case 4:
                    srcX = x;
                    srcY = srcHeight - 1U - y;
                    break;
                case 5:
                    srcX = y;
                    srcY = x;
                    break;
                case 6:
                    srcX = y;
                    srcY = srcHeight - 1U - x;
                    break;
                case 7:
                    srcX = srcWidth - 1U - y;
                    srcY = srcHeight - 1U - x;
                    break;
                case 8:
                    srcX = srcWidth - 1U - y;
                    srcY = x;
                    break;
                case 1:
                default:
                    srcX = x;
                    srcY = y;
                    break;
            }

            const size_t srcIndex =
                (static_cast<size_t>(srcY) * static_cast<size_t>(srcWidth) + srcX) * 4U;
            const size_t dstIndex =
                (static_cast<size_t>(y) * static_cast<size_t>(dstWidth) + x) * 4U;

            output[dstIndex + 0] = image.pixels[srcIndex + 0];
            output[dstIndex + 1] = image.pixels[srcIndex + 1];
            output[dstIndex + 2] = image.pixels[srcIndex + 2];
            output[dstIndex + 3] = image.pixels[srcIndex + 3];
        }
    }

    image.width = dstWidth;
    image.height = dstHeight;
    image.pixels.swap(output);
}

HRESULT GetImageBasicInfo(const std::wstring& path,
                          UINT& width,
                          UINT& height,
                          unsigned long long& fileSize) {
    width = 0;
    height = 0;
    fileSize = 0;

    if (!GetFileSizeBytes(path, fileSize)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    IWICImagingFactory* factory = nullptr;
    IWICBitmapDecoder* decoder = nullptr;
    IWICBitmapFrameDecode* frame = nullptr;

    HRESULT hr = CreateImagingFactory(&factory);
    if (SUCCEEDED(hr)) {
        hr = OpenFirstFrame(factory, path, &decoder, &frame);
    }
    if (SUCCEEDED(hr)) {
        hr = frame->GetSize(&width, &height);
    }
    if (SUCCEEDED(hr)) {
        const USHORT orientation = ReadOrientation(frame);
        if (orientation >= 5 && orientation <= 8) {
            std::swap(width, height);
        }
    }

    SafeRelease(frame);
    SafeRelease(decoder);
    SafeRelease(factory);
    return hr;
}

HRESULT LoadImageData(IWICImagingFactory* factory,
                      const std::wstring& path,
                      ImageData& image,
                      unsigned long long& fileSize) {
    image = {};
    fileSize = 0;

    if (factory == nullptr) {
        return E_POINTER;
    }

    if (!GetFileSizeBytes(path, fileSize)) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    IWICBitmapDecoder* decoder = nullptr;
    IWICBitmapFrameDecode* frame = nullptr;
    IWICFormatConverter* converter = nullptr;

    HRESULT hr = OpenFirstFrame(factory, path, &decoder, &frame);
    if (SUCCEEDED(hr)) {
        hr = frame->GetSize(&image.width, &image.height);
    }
    if (SUCCEEDED(hr)) {
        hr = factory->CreateFormatConverter(&converter);
    }
    if (SUCCEEDED(hr)) {
        hr = converter->Initialize(frame, GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone,
                                   nullptr, 0.0, WICBitmapPaletteTypeCustom);
    }
    if (SUCCEEDED(hr)) {
        const unsigned long long stride64 =
            static_cast<unsigned long long>(image.width) * 4ULL;
        const unsigned long long bytes64 =
            stride64 * static_cast<unsigned long long>(image.height);
        if (stride64 > std::numeric_limits<UINT>::max() ||
            bytes64 > std::numeric_limits<UINT>::max()) {
            hr = E_OUTOFMEMORY;
        } else {
            image.pixels.resize(static_cast<size_t>(bytes64));
            hr = converter->CopyPixels(nullptr, static_cast<UINT>(stride64),
                                       static_cast<UINT>(bytes64), image.pixels.data());
        }
    }
    if (SUCCEEDED(hr)) {
        ApplyOrientation(image, ReadOrientation(frame));
    }

    SafeRelease(converter);
    SafeRelease(frame);
    SafeRelease(decoder);
    return hr;
}

HRESULT ResizeImage(IWICImagingFactory* factory,
                    const ImageData& source,
                    UINT width,
                    UINT height,
                    std::vector<BYTE>& output) {
    if (factory == nullptr) {
        return E_POINTER;
    }
    if (width == 0 || height == 0) {
        return E_INVALIDARG;
    }
    if (source.width == width && source.height == height) {
        output = source.pixels;
        return S_OK;
    }

    if (source.pixels.size() > std::numeric_limits<UINT>::max()) {
        return E_OUTOFMEMORY;
    }

    IWICBitmap* bitmap = nullptr;
    IWICBitmapScaler* scaler = nullptr;

    HRESULT hr = factory->CreateBitmapFromMemory(
        source.width, source.height, GUID_WICPixelFormat32bppBGRA, source.width * 4U,
        static_cast<UINT>(source.pixels.size()), const_cast<BYTE*>(source.pixels.data()), &bitmap);
    if (SUCCEEDED(hr)) {
        hr = factory->CreateBitmapScaler(&scaler);
    }
    if (SUCCEEDED(hr)) {
        hr = scaler->Initialize(bitmap, width, height, WICBitmapInterpolationModeFant);
    }
    if (SUCCEEDED(hr)) {
        const unsigned long long stride64 = static_cast<unsigned long long>(width) * 4ULL;
        const unsigned long long bytes64 = stride64 * static_cast<unsigned long long>(height);
        if (stride64 > std::numeric_limits<UINT>::max() ||
            bytes64 > std::numeric_limits<UINT>::max()) {
            hr = E_OUTOFMEMORY;
        } else {
            output.resize(static_cast<size_t>(bytes64));
            hr = scaler->CopyPixels(nullptr, static_cast<UINT>(stride64),
                                    static_cast<UINT>(bytes64), output.data());
        }
    }

    SafeRelease(scaler);
    SafeRelease(bitmap);
    return hr;
}

void FlattenToBgr24(const std::vector<BYTE>& bgra,
                    UINT width,
                    UINT height,
                    std::vector<BYTE>& bgr24) {
    const size_t pixelCount = static_cast<size_t>(width) * static_cast<size_t>(height);
    bgr24.resize(pixelCount * 3U);

    for (size_t i = 0; i < pixelCount; ++i) {
        const BYTE b = bgra[i * 4U + 0U];
        const BYTE g = bgra[i * 4U + 1U];
        const BYTE r = bgra[i * 4U + 2U];
        const BYTE a = bgra[i * 4U + 3U];

        BYTE outB = b;
        BYTE outG = g;
        BYTE outR = r;
        if (a != 255U) {
            outB = static_cast<BYTE>((static_cast<unsigned int>(b) * a +
                                      255U * (255U - a) + 127U) /
                                     255U);
            outG = static_cast<BYTE>((static_cast<unsigned int>(g) * a +
                                      255U * (255U - a) + 127U) /
                                     255U);
            outR = static_cast<BYTE>((static_cast<unsigned int>(r) * a +
                                      255U * (255U - a) + 127U) /
                                     255U);
        }

        bgr24[i * 3U + 0U] = outB;
        bgr24[i * 3U + 1U] = outG;
        bgr24[i * 3U + 2U] = outR;
    }
}

HRESULT EncodeJpegToMemory(IWICImagingFactory* factory,
                           const std::vector<BYTE>& bgr24,
                           UINT width,
                           UINT height,
                           float imageQuality,
                           std::vector<BYTE>& encodedBytes) {
    if (factory == nullptr) {
        return E_POINTER;
    }

    encodedBytes.clear();

    IStream* stream = nullptr;
    IWICBitmapEncoder* encoder = nullptr;
    IWICBitmapFrameEncode* frame = nullptr;
    IPropertyBag2* properties = nullptr;

    HRESULT hr = CreateStreamOnHGlobal(nullptr, TRUE, &stream);
    if (SUCCEEDED(hr)) {
        hr = factory->CreateEncoder(GUID_ContainerFormatJpeg, nullptr, &encoder);
    }
    if (SUCCEEDED(hr)) {
        hr = encoder->Initialize(stream, WICBitmapEncoderNoCache);
    }
    if (SUCCEEDED(hr)) {
        hr = encoder->CreateNewFrame(&frame, &properties);
    }
    if (SUCCEEDED(hr) && properties != nullptr) {
        PROPBAG2 option = {};
        option.pstrName = const_cast<LPOLESTR>(L"ImageQuality");

        VARIANT qualityValue;
        VariantInit(&qualityValue);
        qualityValue.vt = VT_R4;
        qualityValue.fltVal = imageQuality;
        hr = properties->Write(1, &option, &qualityValue);
        VariantClear(&qualityValue);
    }
    if (SUCCEEDED(hr)) {
        hr = frame->Initialize(properties);
    }
    if (SUCCEEDED(hr)) {
        hr = frame->SetSize(width, height);
    }
    if (SUCCEEDED(hr)) {
        WICPixelFormatGUID pixelFormat = GUID_WICPixelFormat24bppBGR;
        hr = frame->SetPixelFormat(&pixelFormat);
        if (SUCCEEDED(hr) && !IsEqualGUID(pixelFormat, GUID_WICPixelFormat24bppBGR)) {
            hr = E_FAIL;
        }
    }
    if (SUCCEEDED(hr)) {
        const unsigned long long stride64 = static_cast<unsigned long long>(width) * 3ULL;
        const unsigned long long size64 = stride64 * static_cast<unsigned long long>(height);
        if (stride64 > std::numeric_limits<UINT>::max() ||
            size64 > std::numeric_limits<UINT>::max() || size64 > bgr24.size()) {
            hr = E_OUTOFMEMORY;
        } else {
            hr = frame->WritePixels(height, static_cast<UINT>(stride64),
                                    static_cast<UINT>(size64),
                                    const_cast<BYTE*>(bgr24.data()));
        }
    }
    if (SUCCEEDED(hr)) {
        hr = frame->Commit();
    }
    if (SUCCEEDED(hr)) {
        hr = encoder->Commit();
    }
    if (SUCCEEDED(hr)) {
        STATSTG stat = {};
        hr = stream->Stat(&stat, STATFLAG_NONAME);
        if (SUCCEEDED(hr)) {
            const unsigned long long byteCount = stat.cbSize.QuadPart;
            if (byteCount > static_cast<unsigned long long>(std::numeric_limits<ULONG>::max())) {
                hr = E_OUTOFMEMORY;
            } else {
                LARGE_INTEGER start = {};
                hr = stream->Seek(start, STREAM_SEEK_SET, nullptr);
                if (SUCCEEDED(hr)) {
                    encodedBytes.resize(static_cast<size_t>(byteCount));
                    ULONG readBytes = 0;
                    hr = stream->Read(encodedBytes.data(), static_cast<ULONG>(byteCount),
                                      &readBytes);
                    if (SUCCEEDED(hr) && readBytes != byteCount) {
                        hr = E_FAIL;
                    }
                }
            }
        }
    }

    SafeRelease(properties);
    SafeRelease(frame);
    SafeRelease(encoder);
    SafeRelease(stream);
    return hr;
}

bool SaveBufferToFile(const std::wstring& path,
                      const std::vector<BYTE>& bytes,
                      std::wstring& errorMessage) {
    errorMessage.clear();
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        errorMessage = FormatSystemError(GetLastError());
        return false;
    }

    size_t offset = 0;
    while (offset < bytes.size()) {
        const DWORD chunkSize =
            static_cast<DWORD>(std::min<size_t>(bytes.size() - offset, 1024U * 1024U));
        DWORD written = 0;
        if (!WriteFile(file, bytes.data() + offset, chunkSize, &written, nullptr) ||
            written != chunkSize) {
            errorMessage = FormatSystemError(GetLastError());
            CloseHandle(file);
            return false;
        }
        offset += written;
    }

    CloseHandle(file);
    return true;
}

double CandidateScore(const Candidate& candidate) {
    return static_cast<double>(candidate.quality) * 2.0 + candidate.scale * 100.0;
}

HRESULT FindCandidateForScale(IWICImagingFactory* factory,
                              const ImageData& source,
                              UINT width,
                              UINT height,
                              double scale,
                              unsigned long long targetBytes,
                              Candidate& candidate) {
    candidate = {};
    candidate.scale = scale;
    candidate.width = width;
    candidate.height = height;

    std::vector<BYTE> scaledPixels;
    HRESULT hr = ResizeImage(factory, source, width, height, scaledPixels);
    if (FAILED(hr)) {
        return hr;
    }

    std::vector<BYTE> bgr24;
    FlattenToBgr24(scaledPixels, width, height, bgr24);

    std::vector<BYTE> bestBytes;
    hr = EncodeJpegToMemory(factory, bgr24, width, height, 1.0f, bestBytes);
    if (FAILED(hr)) {
        return hr;
    }

    if (bestBytes.size() <= targetBytes) {
        candidate.valid = true;
        candidate.quality = 100;
        candidate.size = bestBytes.size();
        candidate.encodedBytes.swap(bestBytes);
        candidate.score = CandidateScore(candidate);
        return S_OK;
    }

    std::vector<BYTE> lowBytes;
    hr = EncodeJpegToMemory(factory, bgr24, width, height,
                            static_cast<float>(kHardMinQuality) / 100.0f, lowBytes);
    if (FAILED(hr)) {
        return hr;
    }

    if (lowBytes.size() > targetBytes) {
        return S_OK;
    }

    candidate.valid = true;
    candidate.quality = kHardMinQuality;
    candidate.size = lowBytes.size();
    candidate.encodedBytes = lowBytes;

    int low = kHardMinQuality + 1;
    int high = 99;
    while (low <= high) {
        const int mid = low + (high - low) / 2;
        std::vector<BYTE> testBytes;
        hr = EncodeJpegToMemory(factory, bgr24, width, height,
                                static_cast<float>(mid) / 100.0f, testBytes);
        if (FAILED(hr)) {
            return hr;
        }

        if (testBytes.size() <= targetBytes) {
            candidate.quality = mid;
            candidate.size = testBytes.size();
            candidate.encodedBytes.swap(testBytes);
            low = mid + 1;
        } else {
            high = mid - 1;
        }
    }

    candidate.score = CandidateScore(candidate);
    return S_OK;
}

std::vector<double> BuildScalePlan() {
    return {1.00, 0.96, 0.92, 0.88, 0.84, 0.80, 0.76, 0.72, 0.68,
            0.64, 0.60, 0.56, 0.52, 0.48, 0.44, 0.40, 0.36, 0.32,
            0.28, 0.24, 0.20, 0.16, 0.12, 0.09, 0.07, 0.05};
}

CompressionResult CompressImage(const CompressionSettings& settings) {
    CompressionResult result;

    HRESULT comHr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    const bool uninitializeCom = SUCCEEDED(comHr);
    if (FAILED(comHr) && comHr != RPC_E_CHANGED_MODE) {
        result.message = std::wstring(Tr(L"\u521D\u59CB\u5316 COM \u5931\u8D25\uff1A",
                                         L"Failed to initialize COM: ")) +
                         FormatHResult(comHr);
        return result;
    }

    IWICImagingFactory* factory = nullptr;
    ImageData source;

    HRESULT hr = CreateImagingFactory(&factory);
    if (SUCCEEDED(hr)) {
        hr = LoadImageData(factory, settings.inputPath, source, result.inputBytes);
    }

    if (FAILED(hr)) {
        SafeRelease(factory);
        if (uninitializeCom) {
            CoUninitialize();
        }
        result.message = std::wstring(Tr(L"\u8BFB\u53D6\u56FE\u7247\u5931\u8D25\uff1A",
                                         L"Failed to read the image: ")) +
                         FormatHResult(hr);
        return result;
    }

    result.inputWidth = source.width;
    result.inputHeight = source.height;

    Candidate bestCandidate;
    UINT lastWidth = 0;
    UINT lastHeight = 0;

    for (double scale : BuildScalePlan()) {
        const UINT width = std::max<UINT>(
            1U, static_cast<UINT>(std::lround(static_cast<double>(source.width) * scale)));
        const UINT height = std::max<UINT>(
            1U, static_cast<UINT>(std::lround(static_cast<double>(source.height) * scale)));

        if (width == lastWidth && height == lastHeight) {
            continue;
        }
        lastWidth = width;
        lastHeight = height;

        Candidate candidate;
        hr = FindCandidateForScale(factory, source, width, height, scale, settings.targetBytes,
                                   candidate);
        if (FAILED(hr)) {
            SafeRelease(factory);
            if (uninitializeCom) {
                CoUninitialize();
            }
            result.message = std::wstring(Tr(L"\u538B\u7F29\u5931\u8D25\uff1A",
                                             L"Compression failed: ")) +
                             FormatHResult(hr);
            return result;
        }

        if (!candidate.valid) {
            continue;
        }

        if (candidate.quality >= kPreferredQuality) {
            bestCandidate = std::move(candidate);
            break;
        }

        if (!bestCandidate.valid || candidate.score > bestCandidate.score + 0.01) {
            bestCandidate = std::move(candidate);
        }
    }

    if (!bestCandidate.valid) {
        SafeRelease(factory);
        if (uninitializeCom) {
            CoUninitialize();
        }
        result.message = Tr(
            L"\u76EE\u6807\u5927\u5C0F\u8FC7\u5C0F\uff0c\u5373\u4F7F\u5927\u5E45\u964D\u4F4E\u8D28\u91CF"
            L"\u548C\u5206\u8FA8\u7387\u4ECD\u96BE\u4EE5\u8FBE\u5230\u3002\r\n"
            L"\u53EF\u4EE5\u5C1D\u8BD5\u589E\u5927\u76EE\u6807 KB \u6216\u6362\u4E00\u5F20\u56FE\u3002",
            L"The target size is too small. Even with much lower quality and resolution, "
            L"the image still cannot fit.\r\nTry a larger KB target or another image.");
        return result;
    }

    std::wstring saveError;
    if (!SaveBufferToFile(settings.outputPath, bestCandidate.encodedBytes, saveError)) {
        SafeRelease(factory);
        if (uninitializeCom) {
            CoUninitialize();
        }
        result.message = std::wstring(Tr(L"\u5199\u5165\u8F93\u51FA\u6587\u4EF6\u5931\u8D25\uff1A",
                                         L"Failed to write the output file: ")) +
                         saveError;
        return result;
    }

    result.success = true;
    result.targetMet = true;
    result.outputBytes = bestCandidate.size;
    result.outputWidth = bestCandidate.width;
    result.outputHeight = bestCandidate.height;
    result.quality = bestCandidate.quality;
    result.scale = bestCandidate.scale;

    std::wostringstream oss;
    oss << Tr(L"\u538B\u7F29\u5B8C\u6210", L"Compression complete") << L"\r\n"
        << Tr(L"\u8F93\u51FA\u6587\u4EF6\uff1A", L"Output file: ") << settings.outputPath
        << L"\r\n"
        << Tr(L"\u539F\u59CB\u5927\u5C0F\uff1A", L"Original size: ")
        << FormatBytes(result.inputBytes) << L"\r\n"
        << Tr(L"\u8F93\u51FA\u5927\u5C0F\uff1A", L"Output size: ")
        << FormatBytes(result.outputBytes) << L"\r\n"
        << Tr(L"\u539F\u59CB\u5206\u8FA8\u7387\uff1A", L"Original resolution: ")
        << result.inputWidth << L" x " << result.inputHeight << L"\r\n"
        << Tr(L"\u8F93\u51FA\u5206\u8FA8\u7387\uff1A", L"Output resolution: ")
        << result.outputWidth << L" x " << result.outputHeight << L"\r\n"
        << Tr(L"\u6700\u7EC8 JPEG \u8D28\u91CF\uff1A", L"Final JPEG quality: ")
        << result.quality << L"\r\n"
        << Tr(L"\u7F29\u653E\u6BD4\u4F8B\uff1A", L"Scale ratio: ") << FormatPercent(result.scale);

    if (result.scale < 0.999) {
        oss << L"\r\n"
            << Tr(L"\u4E3A\u4E86\u8FBE\u5230\u76EE\u6807\u5927\u5C0F\uff0C\u5DF2\u81EA\u52A8\u7F29\u5C0F"
                     L"\u5206\u8FA8\u7387\u3002",
                  L"The resolution was scaled down automatically to hit the target size.");
    }
    if (result.quality < kPreferredQuality) {
        oss << L"\r\n"
            << Tr(L"\u76EE\u6807\u5927\u5C0F\u8F83\u7D27\uff0C\u5DF2\u5C3D\u91CF\u5728\u5B57\u8282\u9650\u5236"
                     L"\u5185\u4FDD\u7559\u753B\u8D28\u3002",
                  L"The target size is tight, so image quality was preserved as much as "
                  L"possible within the byte limit.");
    }
    result.message = oss.str();

    SafeRelease(factory);
    if (uninitializeCom) {
        CoUninitialize();
    }
    return result;
}

void UpdateInfo(AppState* app, const std::wstring& text) {
    if (app != nullptr && app->infoEdit != nullptr) {
        SetWindowString(app->infoEdit, text);
    }
}

void SetBusy(AppState* app, bool busy) {
    if (app == nullptr) {
        return;
    }

    app->busy = busy;
    EnableWindow(app->inputEdit, !busy);
    EnableWindow(app->inputButton, !busy);
    EnableWindow(app->outputEdit, !busy);
    EnableWindow(app->outputButton, !busy);
    EnableWindow(app->targetEdit, !busy);
    EnableWindow(app->compressButton, !busy);
}

void SyncAutoOutputFromInput(AppState* app) {
    if (app == nullptr) {
        return;
    }

    const std::wstring inputPath = Trimmed(GetWindowString(app->inputEdit));
    const std::wstring currentOutput = Trimmed(GetWindowString(app->outputEdit));
    const std::wstring autoOutput = DefaultOutputPath(inputPath);

    if (currentOutput.empty() || currentOutput == app->lastAutoOutputPath) {
        SetWindowString(app->outputEdit, autoOutput);
    }
    app->lastAutoOutputPath = autoOutput;
}

void RefreshInputInfo(AppState* app) {
    if (app == nullptr) {
        return;
    }

    const std::wstring inputPath = Trimmed(GetWindowString(app->inputEdit));
    if (inputPath.empty()) {
        UpdateInfo(app, Tr(L"\u62D6\u5165\u56FE\u7247\u6216\u70B9\u51FB\u201C\u6D4F\u89C8...\u201D"
                           L"\u9009\u62E9\u6587\u4EF6\u3002",
                           L"Drag an image here or click \"Browse...\" to choose a file."));
        return;
    }

    if (!FileExists(inputPath)) {
        UpdateInfo(app, Tr(L"\u627E\u4E0D\u5230\u8F93\u5165\u6587\u4EF6\u3002\r\n"
                           L"\u8BF7\u68C0\u67E5\u8DEF\u5F84\u662F\u5426\u6B63\u786E\u3002",
                           L"Input file not found.\r\nCheck whether the path is correct."));
        return;
    }

    UINT width = 0;
    UINT height = 0;
    unsigned long long fileSize = 0;
    const HRESULT hr = GetImageBasicInfo(inputPath, width, height, fileSize);
    if (FAILED(hr)) {
        UpdateInfo(app,
                   std::wstring(Tr(L"\u56FE\u7247\u5DF2\u9009\u62E9\uff0c\u4F46\u65E0\u6CD5\u8BFB\u53D6"
                                   L"\u57FA\u672C\u4FE1\u606F\uff1A",
                                   L"The image is selected, but basic info could not be read: ")) +
                       FormatHResult(hr));
        return;
    }

    std::wostringstream oss;
    oss << Tr(L"\u5DF2\u9009\u62E9\uff1A", L"Selected: ") << inputPath << L"\r\n"
        << Tr(L"\u539F\u59CB\u5927\u5C0F\uff1A", L"Original size: ") << FormatBytes(fileSize)
        << L"\r\n"
        << Tr(L"\u5206\u8FA8\u7387\uff1A", L"Resolution: ") << width << L" x " << height
        << L"\r\n"
        << Tr(L"\u8F93\u51FA\u683C\u5F0F\uff1AJPEG", L"Output format: JPEG") << L"\r\n"
        << Tr(L"\u538B\u7F29\u65F6\u4F1A\u4F18\u5148\u901A\u8FC7 JPEG \u8D28\u91CF\u63A7\u5236"
               L"\u5927\u5C0F\uff0c\u5FC5\u8981\u65F6\u81EA\u52A8\u7F29\u5C0F\u5206\u8FA8\u7387\u3002",
               L"Compression prefers adjusting JPEG quality first, and scales the resolution "
               L"down only when necessary.");
    UpdateInfo(app, oss.str());
}

void SetInputPath(AppState* app, const std::wstring& inputPath) {
    if (app == nullptr) {
        return;
    }

    SetWindowString(app->inputEdit, inputPath);
    SyncAutoOutputFromInput(app);
    RefreshInputInfo(app);
}

bool BrowseForInputFile(HWND owner, std::wstring& selectedPath) {
    wchar_t fileBuffer[32768] = {};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = InputFileFilter();
    ofn.lpstrFile = fileBuffer;
    ofn.nMaxFile = static_cast<DWORD>(sizeof(fileBuffer) / sizeof(fileBuffer[0]));
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

    if (!GetOpenFileNameW(&ofn)) {
        return false;
    }

    selectedPath = fileBuffer;
    return true;
}

bool BrowseForOutputFile(HWND owner,
                         const std::wstring& suggestedPath,
                         std::wstring& selectedPath) {
    wchar_t fileBuffer[32768] = {};
    lstrcpynW(fileBuffer, suggestedPath.c_str(),
              static_cast<int>(sizeof(fileBuffer) / sizeof(fileBuffer[0])));

    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = OutputFileFilter();
    ofn.lpstrFile = fileBuffer;
    ofn.nMaxFile = static_cast<DWORD>(sizeof(fileBuffer) / sizeof(fileBuffer[0]));
    ofn.lpstrDefExt = L"jpg";
    ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;

    if (!GetSaveFileNameW(&ofn)) {
        return false;
    }

    selectedPath = ForceJpegExtension(fileBuffer);
    return true;
}

void LayoutControls(AppState* app) {
    if (app == nullptr || app->hwnd == nullptr) {
        return;
    }

    RECT rect = {};
    GetClientRect(app->hwnd, &rect);

    const int width = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    const int margin = 16;
    const int labelWidth = 112;
    const int buttonWidth = 100;
    const int rowHeight = 24;
    const int gap = 10;
    const int rowSpacing = 12;

    const int xLabel = margin;
    const int xEdit = xLabel + labelWidth;
    const int xButton = width - margin - buttonWidth;
    const int editWidth = std::max(140, xButton - gap - xEdit);

    int y = margin;
    MoveWindow(app->inputLabel, xLabel, y + 3, labelWidth, rowHeight, TRUE);
    MoveWindow(app->inputEdit, xEdit, y, editWidth, rowHeight, TRUE);
    MoveWindow(app->inputButton, xButton, y, buttonWidth, rowHeight, TRUE);

    y += rowHeight + rowSpacing;
    MoveWindow(app->outputLabel, xLabel, y + 3, labelWidth, rowHeight, TRUE);
    MoveWindow(app->outputEdit, xEdit, y, editWidth, rowHeight, TRUE);
    MoveWindow(app->outputButton, xButton, y, buttonWidth, rowHeight, TRUE);

    y += rowHeight + rowSpacing;
    MoveWindow(app->targetLabel, xLabel, y + 3, labelWidth, rowHeight, TRUE);
    MoveWindow(app->targetEdit, xEdit, y, 120, rowHeight, TRUE);
    MoveWindow(app->compressButton, width - margin - 128, y - 2, 128, rowHeight + 4, TRUE);

    y += rowHeight + rowSpacing;
    MoveWindow(app->hintStatic, margin, y, width - margin * 2, 56, TRUE);

    y += 62;
    const int infoHeight = std::max(120, height - y - margin);
    MoveWindow(app->infoEdit, margin, y, width - margin * 2, infoHeight, TRUE);
}

void ApplyFont(HWND hwnd, HFONT font) {
    SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
}

void CreateUi(AppState* app) {
    if (app == nullptr || app->hwnd == nullptr) {
        return;
    }

    app->font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));

    app->inputLabel =
        CreateWindowW(L"STATIC", Tr(L"\u8F93\u5165\u56FE\u7247", L"Input Image"),
                      WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, app->hwnd, nullptr, nullptr, nullptr);
    app->inputEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                     WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0,
                                     0, 0, app->hwnd,
                                     reinterpret_cast<HMENU>(IDC_INPUT_EDIT), nullptr, nullptr);
    app->inputButton =
        CreateWindowW(L"BUTTON", Tr(L"\u6D4F\u89C8...", L"Browse..."),
                      WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 0, 0, app->hwnd,
                      reinterpret_cast<HMENU>(IDC_INPUT_BUTTON), nullptr, nullptr);

    app->outputLabel =
        CreateWindowW(L"STATIC", Tr(L"\u8F93\u51FA\u6587\u4EF6", L"Output File"),
                      WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, app->hwnd, nullptr, nullptr, nullptr);
    app->outputEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0,
                                      0, 0, app->hwnd,
                                      reinterpret_cast<HMENU>(IDC_OUTPUT_EDIT), nullptr, nullptr);
    app->outputButton =
        CreateWindowW(L"BUTTON", Tr(L"\u53E6\u5B58\u4E3A...", L"Save As..."),
                      WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 0, 0, app->hwnd,
                      reinterpret_cast<HMENU>(IDC_OUTPUT_BUTTON), nullptr, nullptr);

    app->targetLabel =
        CreateWindowW(L"STATIC", Tr(L"\u76EE\u6807\u5927\u5C0F (KB)", L"Target Size (KB)"),
                      WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, app->hwnd, nullptr, nullptr, nullptr);
    app->targetEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"300",
                                      WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0,
                                      0, 0, app->hwnd,
                                      reinterpret_cast<HMENU>(IDC_TARGET_EDIT), nullptr, nullptr);
    app->compressButton =
        CreateWindowW(L"BUTTON", Tr(L"\u5F00\u59CB\u538B\u7F29", L"Compress"),
                      WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 0, 0, app->hwnd,
                      reinterpret_cast<HMENU>(IDC_COMPRESS_BUTTON), nullptr, nullptr);

    app->hintStatic = CreateWindowW(
        L"STATIC",
        Tr(L"\u8BF4\u660E\uff1A\u8F93\u51FA\u4E3A JPEG\uff0c\u4F18\u5148\u4FDD\u7559\u753B"
           L"\u8D28\uff1B\u82E5\u4ECD\u4E0D\u591F\uff0C\u4F1A\u81EA\u52A8\u7F29\u5C0F\u5206\u8FA8"
           L"\u7387\u3002\r\n\u900F\u660E\u80CC\u666F\u4F1A\u94FA\u767D\u3002",
           L"Output is JPEG. Quality is reduced first; if needed, resolution is scaled down.\r\n"
           L"Transparent areas are flattened to white."),
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, app->hwnd,
        reinterpret_cast<HMENU>(IDC_HINT_STATIC), nullptr, nullptr);

    app->infoEdit =
        CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                        WS_CHILD | WS_VISIBLE | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL |
                            ES_READONLY | WS_VSCROLL,
                        0, 0, 0, 0, app->hwnd, reinterpret_cast<HMENU>(IDC_INFO_EDIT), nullptr,
                        nullptr);

    ApplyFont(app->inputLabel, app->font);
    ApplyFont(app->inputEdit, app->font);
    ApplyFont(app->inputButton, app->font);
    ApplyFont(app->outputLabel, app->font);
    ApplyFont(app->outputEdit, app->font);
    ApplyFont(app->outputButton, app->font);
    ApplyFont(app->targetLabel, app->font);
    ApplyFont(app->targetEdit, app->font);
    ApplyFont(app->compressButton, app->font);
    ApplyFont(app->hintStatic, app->font);
    ApplyFont(app->infoEdit, app->font);

    DragAcceptFiles(app->hwnd, TRUE);
    LayoutControls(app);
    RefreshInputInfo(app);
}

bool CollectSettingsFromUi(AppState* app, CompressionSettings& settings) {
    const std::wstring inputPath = Trimmed(GetWindowString(app->inputEdit));
    if (inputPath.empty()) {
        MessageBoxW(app->hwnd,
                    Tr(L"\u8BF7\u5148\u9009\u62E9\u56FE\u7247\u6587\u4EF6\u3002",
                       L"Please choose an image file first."),
                    NoticeCaption(), MB_OK | MB_ICONWARNING);
        return false;
    }
    if (!FileExists(inputPath)) {
        MessageBoxW(app->hwnd,
                    Tr(L"\u8F93\u5165\u6587\u4EF6\u4E0D\u5B58\u5728\u3002",
                       L"The input file does not exist."),
                    NoticeCaption(), MB_OK | MB_ICONWARNING);
        return false;
    }

    double targetKb = 0.0;
    if (!TryParseDouble(GetWindowString(app->targetEdit), targetKb) || targetKb <= 1.0) {
        MessageBoxW(app->hwnd,
                    Tr(L"\u8BF7\u8F93\u5165\u6709\u6548\u7684\u76EE\u6807\u5927\u5C0F\uff08KB\uff09"
                       L"\u3002",
                       L"Please enter a valid target size in KB."),
                    NoticeCaption(), MB_OK | MB_ICONWARNING);
        return false;
    }

    std::wstring outputPath = Trimmed(GetWindowString(app->outputEdit));
    if (outputPath.empty()) {
        outputPath = DefaultOutputPath(inputPath);
    }
    outputPath = ForceJpegExtension(outputPath);
    if (outputPath.empty()) {
        MessageBoxW(app->hwnd,
                    Tr(L"\u8F93\u51FA\u8DEF\u5F84\u4E0D\u80FD\u4E3A\u7A7A\u3002",
                       L"The output path cannot be empty."),
                    NoticeCaption(), MB_OK | MB_ICONWARNING);
        return false;
    }

    SetWindowString(app->outputEdit, outputPath);

    settings.inputPath = inputPath;
    settings.outputPath = outputPath;
    settings.targetBytes = static_cast<unsigned long long>(std::llround(targetKb * 1024.0));
    return true;
}

DWORD WINAPI CompressionThreadProc(LPVOID parameter) {
    std::unique_ptr<WorkerContext> context(static_cast<WorkerContext*>(parameter));
    std::unique_ptr<CompressionResult> result(
        new CompressionResult(CompressImage(context->settings)));

    if (!PostMessageW(context->hwnd, WM_COMPRESS_DONE, 0,
                      reinterpret_cast<LPARAM>(result.get()))) {
        result.reset();
    } else {
        result.release();
    }
    return 0;
}

void StartCompression(AppState* app) {
    if (app == nullptr || app->busy) {
        return;
    }

    CompressionSettings settings;
    if (!CollectSettingsFromUi(app, settings)) {
        return;
    }

    UpdateInfo(app, Tr(L"\u6B63\u5728\u538B\u7F29\uff0C\u8BF7\u7A0D\u5019...",
                       L"Compressing, please wait..."));
    SetBusy(app, true);

    WorkerContext* context = new WorkerContext;
    context->hwnd = app->hwnd;
    context->settings = settings;

    HANDLE thread = CreateThread(nullptr, 0, CompressionThreadProc, context, 0, nullptr);
    if (thread == nullptr) {
        delete context;
        SetBusy(app, false);
        UpdateInfo(app,
                   std::wstring(Tr(L"\u65E0\u6CD5\u542F\u52A8\u540E\u53F0\u538B\u7F29\u7EBF\u7A0B"
                                   L"\uff1A",
                                   L"Failed to start the background compression thread: ")) +
                       FormatSystemError(GetLastError()));
        return;
    }

    CloseHandle(thread);
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    AppState* app = reinterpret_cast<AppState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (message) {
        case WM_NCCREATE: {
            CREATESTRUCTW* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
            AppState* initialApp = static_cast<AppState*>(create->lpCreateParams);
            initialApp->hwnd = hwnd;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(initialApp));
            return TRUE;
        }

        case WM_CREATE:
            CreateUi(app);
            return 0;

        case WM_SIZE:
            LayoutControls(app);
            return 0;

        case WM_GETMINMAXINFO: {
            MINMAXINFO* info = reinterpret_cast<MINMAXINFO*>(lParam);
            info->ptMinTrackSize.x = 620;
            info->ptMinTrackSize.y = 360;
            return 0;
        }

        case WM_DROPFILES: {
            HDROP drop = reinterpret_cast<HDROP>(wParam);
            wchar_t filePath[32768] = {};
            if (DragQueryFileW(drop, 0, filePath,
                               static_cast<UINT>(sizeof(filePath) / sizeof(filePath[0])))) {
                SetInputPath(app, filePath);
            }
            DragFinish(drop);
            return 0;
        }

        case WM_COMMAND: {
            const int controlId = LOWORD(wParam);
            const int notify = HIWORD(wParam);

            if (controlId == IDC_INPUT_BUTTON && notify == BN_CLICKED) {
                std::wstring selectedPath;
                if (BrowseForInputFile(hwnd, selectedPath)) {
                    SetInputPath(app, selectedPath);
                }
                return 0;
            }

            if (controlId == IDC_OUTPUT_BUTTON && notify == BN_CLICKED) {
                std::wstring selectedPath;
                const std::wstring currentOutput = Trimmed(GetWindowString(app->outputEdit));
                const std::wstring fallbackOutput =
                    currentOutput.empty() ? DefaultOutputPath(Trimmed(GetWindowString(app->inputEdit)))
                                          : currentOutput;
                if (BrowseForOutputFile(hwnd, fallbackOutput, selectedPath)) {
                    SetWindowString(app->outputEdit, selectedPath);
                }
                return 0;
            }

            if (controlId == IDC_COMPRESS_BUTTON && notify == BN_CLICKED) {
                StartCompression(app);
                return 0;
            }

            if (controlId == IDC_INPUT_EDIT &&
                (notify == EN_CHANGE || notify == EN_KILLFOCUS) && !app->busy) {
                SyncAutoOutputFromInput(app);
                if (notify == EN_KILLFOCUS) {
                    RefreshInputInfo(app);
                }
                return 0;
            }

            return 0;
        }

        case WM_COMPRESS_DONE: {
            std::unique_ptr<CompressionResult> result(
                reinterpret_cast<CompressionResult*>(lParam));
            SetBusy(app, false);
            if (result) {
                UpdateInfo(app, result->message);
                if (!result->success) {
                    MessageBoxW(hwnd, result->message.c_str(), CompressionFailedCaption(),
                                MB_OK | MB_ICONERROR);
                }
            }
            return 0;
        }

        case WM_DESTROY:
            PostQuitMessage(0);
            return 0;
    }

    return DefWindowProcW(hwnd, message, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int showCommand) {
    g_appLanguage = DetectAppLanguage();

    const HRESULT comHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool uninitializeCom = SUCCEEDED(comHr);
    if (FAILED(comHr) && comHr != RPC_E_CHANGED_MODE) {
        MessageBoxW(nullptr, Tr(L"\u521D\u59CB\u5316 COM \u5931\u8D25\u3002",
                                L"Failed to initialize COM."),
                    AppTitle(), MB_OK | MB_ICONERROR);
        return 1;
    }

    const wchar_t kClassName[] = L"LightImageCompressorWindowClass";

    WNDCLASSEXW windowClass = {};
    windowClass.cbSize = sizeof(windowClass);
    windowClass.hInstance = instance;
    windowClass.lpfnWndProc = WindowProc;
    windowClass.lpszClassName = kClassName;
    windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    windowClass.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
    windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);

    if (!RegisterClassExW(&windowClass)) {
        if (uninitializeCom) {
            CoUninitialize();
        }
        return 1;
    }

    AppState app;
    HWND hwnd =
        CreateWindowExW(0, kClassName, AppTitle(), WS_OVERLAPPEDWINDOW, CW_USEDEFAULT,
                        CW_USEDEFAULT, 800, 470, nullptr,
                        nullptr, instance, &app);
    if (hwnd == nullptr) {
        if (uninitializeCom) {
            CoUninitialize();
        }
        return 1;
    }

    ShowWindow(hwnd, showCommand);
    UpdateWindow(hwnd);

    MSG message = {};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }

    if (uninitializeCom) {
        CoUninitialize();
    }
    return static_cast<int>(message.wParam);
}
