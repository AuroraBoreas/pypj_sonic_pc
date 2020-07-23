#include "main.h"
#include "cmath"

float round_n(float num, int dec);

// calculate gain value
float DLL_EXPORT calc_gain(float mic_rms, float init_src_rms, float center)
{
    // calculate absolute difference
    float diff = round_n(fabsf(fabs(mic_rms) - fabs(init_src_rms)), 2);
    // calculate absolute gain value
    float gain = round_n(fabsf(center - diff), 2);
    // conditions determine that gain value is positive or negative
    float cond1 = fabsf(mic_rms) - fabsf(init_src_rms);
    float cond2 = diff - center;

    if (cond1 > 0)
    {
        if (cond2 > 0)
            { return -1 * gain; }
        else
            { return gain; }
    }
    else if (cond1 < 0)
    {
        if (cond2 < 0)
            { return -1 * gain + center *2; }
        else
            { return gain + center * 2; }
    }
    else
    { return gain; }
}

// round float to n decimals precision
float round_n(float num, int dec)
{
    double m = (num < 0.0) ? -1.0 : 1.0;   // check if input is negative
    double pwr = pow(10, dec);
    return float(floor((double)num * m * pwr + 0.5) / pwr) * m;
}

extern "C" DLL_EXPORT BOOL APIENTRY DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    switch (fdwReason)
    {
        case DLL_PROCESS_ATTACH:
            // attach to process
            // return FALSE to fail DLL load
            break;

        case DLL_PROCESS_DETACH:
            // detach from process
            break;

        case DLL_THREAD_ATTACH:
            // attach to thread
            break;

        case DLL_THREAD_DETACH:
            // detach from thread
            break;
    }
    return TRUE; // succesful
}
