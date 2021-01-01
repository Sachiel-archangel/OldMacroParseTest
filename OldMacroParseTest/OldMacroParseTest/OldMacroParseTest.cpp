// OldMacroParseTest.cpp : このファイルには 'main' 関数が含まれています。プログラム実行の開始と終了がそこで行われます。
//

#include "pch.h"
#include <coml2api.h>
#include <shlwapi.h>
#include <iostream>
#include <Windows.h>

#pragma comment (lib, "shlwapi.lib")

#define WINAPI __stdcall
char test[1024 * 256];

typedef void(*LPRTLDECOMPRESSBUFFER)(USHORT CompressionFormat, PUCHAR UncompressedBuffer, ULONG UncompressedBufferSize, PUCHAR CompressedBuffer, ULONG CompressedBufferSize, PULONG FinalUncompressedSize);

//NT_RTL_COMPRESS_API NTSTATUS RtlDecompressBuffer(
//	USHORT CompressionFormat,	// COMPRESSION_FORMAT_LZNT1 を指定。
//	PUCHAR UncompressedBuffer,	//解凍されたデータを受信する、呼び出し元が割り当てたバッファー
//	ULONG  UncompressedBufferSize,	//バッファーのサイズ
//	PUCHAR CompressedBuffer,	// 解凍するデータを含むバッファーへのポインター。
//	ULONG  CompressedBufferSize,	// バッファーのサイズ
//	PULONG FinalUncompressedSize	// 解凍されたデータのサイズ
//);

int OutputFile(LPCWSTR pFilename, void *pData, DWORD dwSize)
{
	HANDLE hFile = NULL;
	DWORD dwNumberOfBytesWritten = 0;

	hFile = CreateFile(pFilename, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

	if (hFile == INVALID_HANDLE_VALUE)
		return -1;

	WriteFile(hFile, pData, dwSize, &dwNumberOfBytesWritten, NULL);

	CloseHandle(hFile);

	return 0;

}

int Decompress(void *pMem, DWORD dwSize, LPCWSTR pStreamName)
{
	DWORD dwTargetSize = 0;
	unsigned char *pTarget;

	// パラメータチェック
	if (pMem == NULL || dwSize <= 3)
		return 0;

	// 解凍処理
	ULONG ulFinalUncompressedSize = 0;
	memset(test, 0x00, sizeof(test));

	HMODULE hd = LoadLibrary(L"ntdll.dll");
	FARPROC proc = GetProcAddress(hd, "RtlDecompressBuffer");
	LPRTLDECOMPRESSBUFFER lpRtlDecompressBuffer = reinterpret_cast<LPRTLDECOMPRESSBUFFER>(proc);

	// 頭出し
	// ルール：
	//   "0x01", "0以外", "0xF0でandした結果が0xB0"
	//   "Attribute"は入れるかどうか要考慮
	// 以上を満たした場合、最初の"0x01の次のアドレスがターゲット"
	pTarget = (unsigned char *)pMem;
	dwTargetSize = dwSize;
	DWORD i;

	for (i = 0; i < dwSize - 3; dwTargetSize--, i++)
	{
		if (pTarget[i] != 0x01)
			continue;

		if (pTarget[i + 1] == 0x00)
			continue;

		if (((pTarget[i + 2]) & 0xF0) != 0xB0)
			continue;

		// 追加
		if (pTarget[i + 3] != 0x00)
			continue;


		i++;
		dwTargetSize--;
		break;
	}

	// ヒットしなかった場合はリターン
	if (i + 3 >= dwSize)
		return 0;

	lpRtlDecompressBuffer(COMPRESSION_FORMAT_LZNT1, (PUCHAR)test, 1024 * 64, (PUCHAR)&pTarget[i], dwTargetSize, &ulFinalUncompressedSize);

	printf("------ Decompress Result ------\r\n");
	printf("%s\r\n", test);
	printf("-------------------------------\r\n\r\n");

	// ファイルに出力
	int iStreamNameLen = 0;
	int iBufferSize = 0;
	LPCWSTR pOutFileName = NULL;

	iStreamNameLen = wcslen(pStreamName);
	iBufferSize = (iStreamNameLen + 13) * 2 + 2;
	pOutFileName = (LPCWSTR)VirtualAlloc(NULL, iBufferSize, MEM_COMMIT, PAGE_READWRITE);
	wcscpy_s((wchar_t *)pOutFileName, (rsize_t)iBufferSize, (wchar_t *)pStreamName);
	wcscat_s((wchar_t *)pOutFileName, (rsize_t)iBufferSize, L"_decompressed");

	OutputFile(pOutFileName, test, ulFinalUncompressedSize);

	//FreeLibrary(hd);

	return 0;
}

int ReadStorage(LPSTORAGE pStg)
{
	HRESULT objHResult = NULL;

	LPMALLOC pMalloc = NULL;	// statstgを解放するため
	HRESULT hr = CoGetMalloc(MEMCTX_TASK, &pMalloc);

	LPENUMSTATSTG pEnum = NULL;
	objHResult = pStg->EnumElements(0, NULL, 0, &pEnum);
	if (FAILED(objHResult)) {
		printf("EnumElements Error.\r\n");
		return 0;
	}

	STATSTG statstg;
	while (pEnum->Next(1, &statstg, NULL) == S_OK) {
		switch (statstg.type) {
		case STGTY_STORAGE:
			LPSTORAGE pSubStg;
			pSubStg = NULL;
			hr = pStg->OpenStorage(statstg.pwcsName,
				NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pSubStg);
			wprintf(L" Storage = %s\n", statstg.pwcsName);
			// wprintf(L"%d Storage = %s\n", nIndent, statstg.pwcsName);
			ReadStorage(pSubStg);
			pSubStg->Release();
			break;
		case STGTY_STREAM:
			LPCWSTR pStreamName = NULL;

			if (((char *)statstg.pwcsName)[0] == 0x01) {
				// 注：statstg.pwcsNameがLPWSTRのため、1加算で2バイト移動。
				pStreamName = statstg.pwcsName + 1;
			}
			else{
				pStreamName = statstg.pwcsName;
			}


			wprintf(L"Stream = %s %llu\n", pStreamName, statstg.cbSize.QuadPart);
			//			wprintf(L"%d Stream = %s %llu\n", nIndent, statstg.pwcsName, statstg.cbSize.QuadPart);
			//			WriteStream(pStg, statstg);
			IStream *pStream;
			ULARGE_INTEGER uiSize;
			void *pMem;
			ULONG uRead;

			pStg->OpenStream(statstg.pwcsName, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, &pStream);
			if (pStream) {
				IStream_Size(pStream, &uiSize);
				pMem = VirtualAlloc(NULL, (SIZE_T)uiSize.LowPart, MEM_COMMIT, PAGE_READWRITE);
				pStream->Read(pMem, uiSize.LowPart, &uRead);
				wprintf(L"    Storage Size = %d\n", uiSize.LowPart);

				// ファイルに出力（バイナリデータをそのまま出力する）
				OutputFile(pStreamName, pMem, uiSize.LowPart);

				// 解凍処理
				Decompress(pMem, uiSize.LowPart, statstg.pwcsName);

				// 後処理
				VirtualFree(pMem, 0, MEM_RELEASE);
				pStream->Release();
			}

			break;
		}
		//pMalloc->Free(statstg);	// メモリリークを避ける
		pMalloc->Free(statstg.pwcsName);	// メモリリークを避ける
	}


	pMalloc->Release();
	pEnum->Release();

	return 0;
}

int main(int argc, LPTSTR *argv[])
{
	HRESULT objHResult = NULL;
	LPSTORAGE pStgRoot = NULL;

	if (argc != 2) {
		printf("Parameter Error.\r\n");
		return 0;
	}

	WCHAR chrWideParam[1000];
	size_t szReturnValue;

	memset(chrWideParam, 0x00, sizeof(chrWideParam));
	mbstowcs_s(&szReturnValue, (wchar_t *)chrWideParam, sizeof(chrWideParam) / 2 - 2, (const char*)argv[1], _TRUNCATE);

	objHResult = StgOpenStorage((CONST WCHAR*)chrWideParam, NULL, STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pStgRoot);
	if (FAILED(objHResult)) {
		printf("StgOpenStorage Error.\r\n");
		return 0;
	}

	ReadStorage(pStgRoot);

	pStgRoot->Release();

	CoUninitialize();
}

// プログラムの実行: Ctrl + F5 または [デバッグ] > [デバッグなしで開始] メニュー
// プログラムのデバッグ: F5 または [デバッグ] > [デバッグの開始] メニュー

// 作業を開始するためのヒント: 
//    1. ソリューション エクスプローラー ウィンドウを使用してファイルを追加/管理します 
//   2. チーム エクスプローラー ウィンドウを使用してソース管理に接続します
//   3. 出力ウィンドウを使用して、ビルド出力とその他のメッセージを表示します
//   4. エラー一覧ウィンドウを使用してエラーを表示します
//   5. [プロジェクト] > [新しい項目の追加] と移動して新しいコード ファイルを作成するか、[プロジェクト] > [既存の項目の追加] と移動して既存のコード ファイルをプロジェクトに追加します
//   6. 後ほどこのプロジェクトを再び開く場合、[ファイル] > [開く] > [プロジェクト] と移動して .sln ファイルを選択します
