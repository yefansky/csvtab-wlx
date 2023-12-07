#include "stdafx.h"

BOOL IsEOL(TCHAR c) 
{
	return c == TEXT('\r') || c == TEXT('\n');
}

BOOL IsNumber(TCHAR* pszVal)
{
	BOOL	bResult = TRUE;
	int		nLen	= (int)_tcslen(pszVal);
	int		nCount	= 0;

	for (int i = 0; bResult && i < nLen; i++) 
	{
		nCount += pszVal[i] == TEXT('.');
		bResult = _istdigit(pszVal[i]) || pszVal[i] == TEXT('.');
	}
	return bResult && nCount < 2;
}

void MergeSortJoiner(int nIndexes[], void* pvData, int l, int m, int r, BOOL bisBackward, BOOL bisNums) 
{
	int		n1	= m - l + 1;
	int		n2	= r - m;
	int*	L	= (int*)calloc(n1, sizeof(int));
	int*	R	= (int*)calloc(n2, sizeof(int));

	for (int i = 0; i < n1; i++)
		L[i] = nIndexes[l + i];
	for (int j = 0; j < n2; j++)
		R[j] = nIndexes[m + 1 + j];

	int i = 0, j = 0, k = l;
	while (i < n1 && j < n2) 
	{
		int cmp = bisNums ? ((double*)pvData)[L[i]] <= ((double*)pvData)[R[j]] : _tcscmp(((TCHAR**)pvData)[L[i]], ((TCHAR**)pvData)[R[j]]) <= 0;
		if (bisBackward)
			cmp = !cmp;

		if (cmp) 
		{
			nIndexes[k] = L[i];
			i++;
		}
		else 
		{
			nIndexes[k] = R[j];
			j++;
		}
		k++;
	}

	while (i < n1) 
	{
		nIndexes[k] = L[i];
		i++;
		k++;
	}

	while (j < n2) 
	{
		nIndexes[k] = R[j];
		j++;
		k++;
	}

	free(L);
	free(R);
}

void MergeSort(int nIndexes[], void* pvData, int l, int r, BOOL bisBackward, BOOL bisNums) 
{
	if (l < r) 
	{
		int m = l + (r - l) / 2;
		MergeSort(nIndexes, pvData, l, m, bisBackward, bisNums);
		MergeSort(nIndexes, pvData, m + 1, r, bisBackward, bisNums);
		MergeSortJoiner(nIndexes, pvData, l, m, r, bisBackward, bisNums);
	}
}
