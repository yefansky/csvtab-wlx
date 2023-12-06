#include "stdafx.h"

BOOL isEOL(TCHAR c) {
	return c == TEXT('\r') || c == TEXT('\n');
}

BOOL isNumber(TCHAR* val) {
	int len = _tcslen(val);
	BOOL res = TRUE;
	int pCount = 0;
	for (int i = 0; res && i < len; i++) {
		pCount += val[i] == TEXT('.');
		res = _istdigit(val[i]) || val[i] == TEXT('.');
	}
	return res && pCount < 2;
}

void mergeSortJoiner(int indexes[], void* data, int l, int m, int r, BOOL isBackward, BOOL isNums) {
	int n1 = m - l + 1;
	int n2 = r - m;

	int* L = (int*)calloc(n1, sizeof(int));
	int* R = (int*)calloc(n2, sizeof(int));

	for (int i = 0; i < n1; i++)
		L[i] = indexes[l + i];
	for (int j = 0; j < n2; j++)
		R[j] = indexes[m + 1 + j];

	int i = 0, j = 0, k = l;
	while (i < n1 && j < n2) {
		int cmp = isNums ? ((double*)data)[L[i]] <= ((double*)data)[R[j]] : _tcscmp(((TCHAR**)data)[L[i]], ((TCHAR**)data)[R[j]]) <= 0;
		if (isBackward)
			cmp = !cmp;

		if (cmp) {
			indexes[k] = L[i];
			i++;
		}
		else {
			indexes[k] = R[j];
			j++;
		}
		k++;
	}

	while (i < n1) {
		indexes[k] = L[i];
		i++;
		k++;
	}

	while (j < n2) {
		indexes[k] = R[j];
		j++;
		k++;
	}

	free(L);
	free(R);
}

void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums) {
	if (l < r) {
		int m = l + (r - l) / 2;
		mergeSort(indexes, data, l, m, isBackward, isNums);
		mergeSort(indexes, data, m + 1, r, isBackward, isNums);
		mergeSortJoiner(indexes, data, l, m, r, isBackward, isNums);
	}
}
