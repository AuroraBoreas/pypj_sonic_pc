#ifndef __MAIN_H__
#define __MAIN_H__

#include <windows.h>
/*  To use this exported function of dll, include this header
 *  in your project.
 */

#ifdef BUILD_DLL
    #define DLL_EXPORT __declspec(dllexport)
#else
    #define DLL_EXPORT __declspec(dllimport)
#endif


extern "C"
{
    float DLL_EXPORT calc_gain(float mic_rms, float init_src_rms, float center);
}
#endif // __MAIN_H__
