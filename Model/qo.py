import numpy as np

def qo_f(pwf,pr,Qmax):
    return (Qmax*(1-0.2*(pwf/pr)-0.8*((pwf/pr)**2)))