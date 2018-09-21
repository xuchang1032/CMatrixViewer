#include <iostream>
#include "CGnssEpoch.h"

int var[6] = {1, 2, 2, 5};

int KaTest(double Inno[], int n)
{
	double sum = 0, thr = var[n];
	for (int i = 0; i < n; i++)
	{
		sum += Inno[i]*Inno[i];
	}
	if (sum > thr)
		return 1;
	else
		return 0;

}

void main(int argc, char *argv)
{
	double Inno[] = {1.25, 10., 19.1, 0.291};
	int kaRes = KaTest(Inno, 4);
}

