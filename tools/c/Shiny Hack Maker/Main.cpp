// Shiny Hack Maker {Special C Version}
// By ZodiacDaGreat
// Special Thanks To interdpth & cearn
/////////////////////////////////////////

// Libraries
#include <windows.h>
#include <stdio.h>
#include <commctrl.h>
#include "resource.h"

// Global Variables
FILE *fp;

unsigned long IWRAMOffset;
unsigned long ShinyHeaderOffset;
unsigned long RNG;
unsigned long Offset;

char EnlargedROM;
char sFilePath[MAX_PATH] = "";
char sHeader[4] = "";

const unsigned char OriginalBytes[32] = {0x39,0x60,0xA,0xE2,0x21,0x78,0x60,0x78,0x0,0x2,0x9,0x18,0xA0,0x78,0x0,0x4,0x9,0x18,0xE0,0x78,0x0,0x6,0x9,0x18,0x79,0x60,0xFE,0xE1,0x0,0x22,0x3B,0x1C};
const unsigned char PatchingBytes[32] = {0x39,0x60,0xA,0xE2,0x2,0xB4,0x2,0x49,0x0,0xF0,0x0,0xF8,0xF,0x47,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0xFE,0xE1,0x0,0x22,0x3B,0x1C};
const unsigned char ShinyHack[94] = {0x2,0xBC,0x2D,0xB5,0x4,0x99,0x8,0x31,0x4,0x91,0x13,0x48,0x0,0xF0,0x26,0xF8,0x5,0x4,0x45,0x19,0xF,0x4A,0x11,0x68,0x0,0x23,0x13,0x60,0x1,0x29,0xB,0xD0,0x21,0x78,0x60,0x78,0x0,0x2,0x9,0x18,0xA0,0x78,0x0,0x4,0x9,0x18,0xE0,0x78,0x0,0x6,0x9,0x18,0x79,0x60,0xC,0xE0,0x21,0x78,0x60,0x78,0x0,0x2,0x9,0x18,0xA0,0x78,0x0,0x4,0x9,0x18,0xE0,0x78,0x0,0x6,0x9,0x18,0x79,0x60,0x69,0x40,0x39,0x60,0x2D,0xBD,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x0,0x47};
const unsigned char FlagRoutine[16] = {0x7,0xB5,0x3,0x48,0x1,0x68,0x1,0x22,0x51,0x40,0x1,0x60,0x7,0xBD,0x0,0x0};

const char *HeaderList[24] = {"AXVE", "AXPE", "BPRE", "BPGE", "BPEE", "BPEF"};

// Functions
void CheckShinyHack(HWND hwnd)
{
	char OriginalByte;
	unsigned long ExtendedOffset;
	char TextExtendedOffset[7] = "";

	fp = fopen(sFilePath, "r+b");
	fseek(fp, ShinyHeaderOffset + 19, SEEK_SET);
	OriginalByte = fgetc(fp);

	if (OriginalByte == 0x78)
	{
		SetWindowText(GetDlgItem(hwnd, lblState), "Not Patched");
		SetWindowText(GetDlgItem(hwnd, txtShinyHack), "");
	}
	else
	{
		SetWindowText(GetDlgItem(hwnd, lblState), "Patched");
		fseek(fp, ShinyHeaderOffset + 16, SEEK_SET);
		fread(&ExtendedOffset, sizeof(int), 1, fp);
		ExtendedOffset = ExtendedOffset & 0xFFFFFF;

		if (OriginalByte == 0x9)
		{
			ExtendedOffset = ExtendedOffset + 0x1000000;
		}
		
		ExtendedOffset = ExtendedOffset - 1;
		sprintf(TextExtendedOffset, "%x", ExtendedOffset);
		SetWindowText(GetDlgItem(hwnd, txtShinyHack), TextExtendedOffset);
		EnableWindow(GetDlgItem(hwnd, txtShinyHack), 0); // Disable Shiny Hack textbox
		EnableWindow(GetDlgItem(hwnd, cmdPatch), 0); // Disable command button Patch
	}

	fclose(fp);
}

void ExtendRoutine(HWND hwnd)
{
	unsigned long Temp;
	char Bank;
	char TextOffset[7] = "";

	Bank = 0x8;
	GetWindowText(GetDlgItem(hwnd, txtShinyHack), TextOffset, 8);
	sscanf(TextOffset, "%x", &Temp);
	if (Temp > 0xFFFFFF)
	{
		Bank = 0x9;
	}

	fp = fopen(sFilePath, "r+b");
	fseek(fp, ShinyHeaderOffset, SEEK_SET);
	fwrite(&PatchingBytes, 32, 1, fp);
	GetWindowText(GetDlgItem(hwnd, txtShinyHack), TextOffset, 8);
	sscanf(TextOffset, "%x", &Offset);
	Offset = Offset + 1;
	fseek(fp, ShinyHeaderOffset + 16, SEEK_SET);
	fwrite(&Offset, 4, 1, fp);
	fseek(fp, ShinyHeaderOffset + 19, SEEK_SET);
	fputc(Bank, fp);
	fclose(fp);
}

void ImplementShinyHack(HWND hwnd)
{
	fp = fopen(sFilePath, "r+b");
	fseek(fp, Offset, SEEK_SET);
	fwrite(&ShinyHack, 94, 1, fp);
	fseek(fp, Offset + 84, SEEK_SET);
	fwrite(&IWRAMOffset, 4, 1, fp);
	fseek(fp, Offset + 88, SEEK_SET);
	fwrite(&RNG, 4, 1, fp);
	fclose(fp);
	MessageBox(hwnd, "Shiny Header is patched.\nThe extended routine is inserted successfully.", "Shiny Hack Maker", MB_ICONINFORMATION);
	CheckShinyHack(hwnd);
}

void InsertFlagRoutine(HWND hwnd)
{
	char txtFlagOffset[7] = "";
	unsigned long FlagOffset;

	GetWindowText(GetDlgItem(hwnd, txtFlagRoutine), txtFlagOffset, 8);
	sscanf(txtFlagOffset, "%x", &FlagOffset); // %u = dec, %x = hex
	fp = fopen(sFilePath, "r+b");
	fseek(fp, FlagOffset, SEEK_SET);
	fwrite(&FlagRoutine, 16, 1, fp);
	fseek(fp, FlagOffset + 16, SEEK_SET);
	fwrite(&IWRAMOffset, 4, 1, fp);
	fclose(fp);
	MessageBox(hwnd, "The flag routine is successfully inserted.", "Shiny Hack Maker", MB_ICONINFORMATION);
}

int GetHeaderIndex(const char *header, const char *headerlist[], int headeramount)
{
	int i;

	for(i=0; i < headeramount; i++)
	if( memcmp(header, headerlist[i], 4) == 0)
	{
	return i;
	}
	return -1;
}

void CheckOptShinyHack(HWND hwnd)
{
	SendMessage(GetDlgItem(hwnd, optShinyHack), BM_SETCHECK, 1, 0); // Check Shiny Hack
	SendMessage(GetDlgItem(hwnd, optFlagRoutine), BM_SETCHECK, 0, 0); // Uncheck Flag Routine
	
	ShowWindow(GetDlgItem(hwnd, fraShinyHack), 1); // Show Shiny Hack frame and contents
	ShowWindow(GetDlgItem(hwnd, txtShinyHack), 1);
	ShowWindow(GetDlgItem(hwnd, cmdPatch), 1);

	ShowWindow(GetDlgItem(hwnd, fraFlagRoutine), 0); // Hide Flag Routine frame and contents
	ShowWindow(GetDlgItem(hwnd, txtFlagRoutine), 0);
	ShowWindow(GetDlgItem(hwnd, cmdInsert), 0);

	SetWindowText(GetDlgItem(hwnd, lblDec2), "94 bytes");
	SetWindowText(GetDlgItem(hwnd, lblHex2), "0x5E bytes");
}

void CheckOptFlagRoutine(HWND hwnd)
{
	SendMessage(GetDlgItem(hwnd, optShinyHack), BM_SETCHECK, 0, 0); // Uncheck Shiny Hack
	SendMessage(GetDlgItem(hwnd, optFlagRoutine), BM_SETCHECK, 1, 0); // Check Flag Routine

	ShowWindow(GetDlgItem(hwnd, fraShinyHack), 0); // Show Flag Routine frame and contents
	ShowWindow(GetDlgItem(hwnd, txtShinyHack), 0);
	ShowWindow(GetDlgItem(hwnd, cmdPatch), 0);

	ShowWindow(GetDlgItem(hwnd, fraFlagRoutine), 1); // Hide Shiny Hack frame and contents
	ShowWindow(GetDlgItem(hwnd, txtFlagRoutine), 1);
	ShowWindow(GetDlgItem(hwnd, cmdInsert), 1);

	SetWindowText(GetDlgItem(hwnd, lblDec2), "32 bytes");
	SetWindowText(GetDlgItem(hwnd, lblHex2), "0x20 bytes");
}

void CheckROMSize(HWND hwnd)
{
	unsigned long ROMSize;

	fp = fopen(sFilePath, "r+b");
	fseek(fp, -1, SEEK_END);
	ROMSize = ftell(fp);
	if (ROMSize > 0xFFFFFF)
	{
		SendMessage(GetDlgItem(hwnd, txtShinyHack), EM_SETLIMITTEXT, 7, 1); // Setting Max Length to 7
		SendMessage(GetDlgItem(hwnd, txtFlagRoutine), EM_SETLIMITTEXT, 7, 1);
		EnlargedROM = 1;
	}
	else
	{
		SendMessage(GetDlgItem(hwnd, txtShinyHack), EM_SETLIMITTEXT, 6, 1); // Setting Max Length to 7
		SendMessage(GetDlgItem(hwnd, txtFlagRoutine), EM_SETLIMITTEXT, 6, 1);
		EnlargedROM = 0;
	}

	fclose (fp);
}

void LoadROM(HWND hwnd)
{
	OPENFILENAME ofn;
	char unsupported[99] = "Error 1: Non-POKéMON ROM unsupported.\nIf it is a POKéMON ROM then you must\nrequest for support :P";

    ZeroMemory(&ofn, sizeof(ofn));

    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
	ofn.lpstrTitle = "Open ROM...";
    ofn.lpstrFilter = "Gameboy Advance ROMs (*.gba,*.agb,*.bin)\0*.gba;*.agb;*.bin";
    ofn.lpstrFile = sFilePath;
    ofn.nMaxFile = MAX_PATH;

    if(GetOpenFileName(&ofn))
    {
		fp = fopen(sFilePath, "r+b");
		fseek(fp, 0xAC, SEEK_SET);
		fread(sHeader, sizeof(int), 1, fp);
		fclose(fp);

		if(GetHeaderIndex(sHeader, HeaderList, 6) == -1)
		{
			MessageBox(NULL, unsupported, "Unsupported ROM", MB_ICONERROR);
		}

		// Ruby AXVE
		if(memcmp(sHeader, "AXVE", 4) == 0)
		{
			SetWindowText(GetDlgItem(hwnd, lblROM), "Pokémon Ruby");
			IWRAMOffset = 0x2042000;
			ShinyHeaderOffset = 0x3D4D8;
			RNG = 0x8040E85;
			
		}
		
		// Sapphire AXPE
		if(memcmp(sHeader, "AXPE", 4) == 0)
		{
			SetWindowText(GetDlgItem(hwnd, lblROM), "Pokémon Sapphire");
			IWRAMOffset = 0x2042000;
			ShinyHeaderOffset = 0x3D4D8;
			RNG = 0x8040E85;
		}

		// Fire Red BPRE
		if(memcmp(sHeader, "BPRE", 4) == 0)
		{
			SetWindowText(GetDlgItem(hwnd, lblROM), "Pokémon Fire Red");
			IWRAMOffset = 0x2022000;
			ShinyHeaderOffset = 0x406C0;
			RNG = 0x8044EC9;
		}

		// Leaf Green BPGE
		if(memcmp(sHeader, "BPGE", 4) == 0)
		{
			SetWindowText(GetDlgItem(hwnd, lblROM), "Pokémon Leaf Green");
			IWRAMOffset = 0x2022000;
			ShinyHeaderOffset = 0x406C0;
			RNG = 0x8044EC9;
		}

		// Emerald BPEE
		if(memcmp(sHeader, "BPEE", 4) == 0)
		{
			SetWindowText(GetDlgItem(hwnd, lblROM), "Pokémon Emerald");
			IWRAMOffset = 0x2042000;
			ShinyHeaderOffset = 0x6AF8C;
			RNG = 0x806F5CD;
		}

		// Emeraude BPEF
		if(memcmp(sHeader, "BPEF", 4) == 0)
		{
			SetWindowText(GetDlgItem(hwnd, lblROM), "Pokémon Emeraude");
			IWRAMOffset = 0x2042000;
			ShinyHeaderOffset = 0x6AF8C;
			RNG = 0x806F5C9;
		}

		SetWindowText(GetDlgItem(hwnd, lblCode), sHeader);
		EnableWindow(GetDlgItem(hwnd, cmdPatch), 1);
		EnableWindow(GetDlgItem(hwnd, txtShinyHack), 1);
		EnableWindow(GetDlgItem(hwnd, cmdInsert), 1);
		EnableWindow(GetDlgItem(hwnd, txtFlagRoutine), 1);
		CheckROMSize(hwnd);
		CheckShinyHack(hwnd);

    }
}

BOOL CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	static HICON hIcon;
	static HICON hIconSm;

	switch(message)
	{
		case WM_INITDIALOG: // Similar to Form Load
			// Loading Icon to Dialog
			hIconSm = (HICON)LoadImage(GetModuleHandle(NULL), MAKEINTRESOURCE(dlgIcon), IMAGE_ICON, 16, 16, 0);
			SendMessage(hwnd, WM_SETICON, ICON_SMALL, (LPARAM)hIconSm);

			CheckOptShinyHack(hwnd);
			break;

		case BN_CLICKED:
			break;

		case BN_DBLCLK:
			break;

		case WM_COMMAND:
			switch(LOWORD(wParam))
			{
				case cmdPatch:
					ExtendRoutine(hwnd);
					ImplementShinyHack(hwnd);
					break;

				case cmdInsert:
					InsertFlagRoutine(hwnd);
					break;
					
				case cmdOkay:
					EndDialog(hwnd, 0);
					break;

				case mnuReadme:
					
					break;

				case mnuAbout:
					DialogBox(GetModuleHandle(NULL), MAKEINTRESOURCE(frmAbout), hwnd, DialogProc);
					break;

				case mnuOpen:
					LoadROM(hwnd);
					break;

				case mnuQuit:
					EndDialog(hwnd,0);
					break;

				case optShinyHack:
					CheckOptShinyHack(hwnd);
					break;

				case optFlagRoutine:
					CheckOptFlagRoutine(hwnd);
					break;

				case IDCANCEL:
					EndDialog(hwnd,0);
					break;
			}

			break;

		case WM_LBUTTONDBLCLK:
			break;

		case WM_DESTROY:
			DestroyIcon(hIcon);
			DestroyIcon(hIconSm);
			break;
	}

	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
	InitCommonControls; // Needed for XP Style	
	DialogBox(hInstance, MAKEINTRESOURCE(frmMain), 0, DialogProc);

	return 0;
}