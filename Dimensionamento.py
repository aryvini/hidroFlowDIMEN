import xlwings as xw
import numpy as np
##remember to add the python lib/bin to the environmental windows path variables
##in order to correctly import numpy into xlwings
from math import pow,sin,cos,acos,sqrt,pi
from scipy.optimize import fsolve
#function to integrated and run inside excel
#calculated an interative implicit equation

##function names pattern:
##hidro_firstSecondTHIRD

@xw.func
@xw.arg('flow', doc="Flow [m³/s]")
@xw.arg('diameter', doc="Diameter in m")
@xw.arg('slope', doc="slope m/m")
@xw.arg('Ks', doc="Manning coefficient, default=100")
@xw.arg('guess', doc="First guess to start the root iteraction, default=0.2")
def hidro_findTheta(flow, diameter, slope, Ks = 110, guess = 0.2):
    """Function to iteractive find a root of the theta expression
    Input all in SI units
    Parameters:
    flow: m³/s
    diameter: m
    slope: m/m
    Ks: Rough coefficient, default=110 (PVC pipes)
    guess: First guess to start the root iteraction, default=0.2
    ------------------------------------------------------------
    Output:
    angle in radians
    """
    if guess is 0:
        return print("Argument error")

    func = lambda val: ((sin(val) + 6.063 * pow(flow/(Ks*sqrt(slope)),0.6)*pow(diameter,-1.6)*pow(val,0.4)) - val)
    try:
        result = fsolve(func,guess)
    except Exception: #If an error occours (math error or something else) change the guess and run the function again recursion
        newguess = guess+0.2
        print ("The root can't be found, trying a different starting guess")
        print ("{}".format(newguess))
        return hidro_findTheta(flow, diameter, slope, Ks = 100, guess = newguess)

    if result[0] >= 2*pi: ##function domain [0,2pi]
        return "Theta is higher than 2*pi, paramters error"

    return result[0]


@xw.func
@xw.arg('flowMax', doc="[m³/s]")
@xw.arg('pipeUsage', doc="Percentage of allowable area to flow, default 0.5")
@xw.arg('slopeMax', doc="Maximum allowed slope of the pipe, default 15/100")
@xw.arg('Ks', doc="Manning coefficient, default=100")

def hidro_minimumDiameter(flowMax,pipeUsage=0.5,slopeMax=15/100,Ks=110):
    """Calculates the minimum possible diameter according to maximum slope, flow and pipe usage
    Input all in SI units
    Parameters:
    flowMax[m³/s]: Maximum flow at the end of the project horizon
    pipeUsage[0 to 1]: Percentage of allowable area to flow default 0.5
    slopeMax[m/m] = Maximum allowed slope of the pipe, default 15/100
    Ks: Rough coefficient, default=110 (PVC pipes)
    ------------------------------------------------------------
    Output:
    Minimum pipe diameter [milimeters]
    """

    ##Input checks
    if pipeUsage<=0:
        return "pipeUsage can't assume 0 neither negative value"
    if pipeUsage>1:
        return "pipeUsage can't be higher than 1"

    if slopeMax<0:
        return "Slope can't assume negative value"
    if slopeMax>1:
        return "Slope can't be higher than 1"


    try:
        theta = 2*acos(1-2*pipeUsage)
        diameter = ((20.159 * flowMax/(Ks*sqrt(slopeMax)))**(3/8))*(theta**(1/4)/(theta-sin(theta))**(5/8))

    except Exception:
        print("Error during the calculation, modify the parameters and try again")
    
        
    return (diameter*1000)



