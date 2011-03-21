#ifndef CALC_H
#define CALC_H
#include <string>
namespace Calc
{
    enum 
    {
        E = 4,
        F = 5,
        G = 6,
        H = 7,
        I = 8,
        L = 11,
        M = 12,
        P = 15,
        Q = 16,
        T = 19
    };
    struct Result
    {
        double E;
        double F;
        double GH;
        double IL;
        double MP;
        double QT;
    };

    Result* calc(double,double,string);
}
#endif
