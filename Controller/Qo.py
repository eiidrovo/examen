import xlwings as xw
import Model.qo as qo
import numpy as np
import matplotlib.pyplot as plt

def main():
    wb = xw.Book.caller()
    mainsheet = wb.sheets["Plot"]
    Pr=mainsheet.range("PR").value
    Qmax=mainsheet.range("Qmax").value
    Observations=int(mainsheet.range("Observations").value)

    PWF = np.linspace(Pr, 0, Observations)
    QO_values= qo.qo_f(PWF,Pr,Qmax)

    sheet_Values=wb.sheets["values"]
    sheet_Values["A2"].options(np.array, transpose=True).value=PWF

    sheet_Values["B2"].options(np.array, transpose=True).value=QO_values
    fig, ax = plt.subplots()
    ax.set_title(f"Qo vs PWF")
    ax.plot(QO_values, PWF, label='Qo')

    ax.legend()
    ax.set_xlabel('Qo [stb/d]')
    ax.set_ylabel('Pressure [psi]')
    ax.grid()

    mainsheet.pictures.add(fig, name='Plot', update=True,
                         left=mainsheet.range('B8').left, top=mainsheet.range('B8').top)

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("Qo.xlsm").set_mock_caller()
    main()
