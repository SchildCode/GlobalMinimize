# GlobalMinimize
### Features
- A Microsoft Excel macro with state-of-the art multi-parameter bounded optimization algorithms that outperform the Excel Solver AddIn when the solution domain contains multiple local minima, discontinuous gradient, or regions with constant cost (no gradient). 
- The application contains improved versions of the following two optimizers: 
  1. An improved version of the Downhill Simplex algorithm by Nelder-Mead (abbreviated "NM"). The original NM algorithm is known to suffer from stagnation and premature convergence, but this improved version uses successive-approximation (SA-NM) with an extra pseudo-expansion point to guarantee convergence to minima. My own improvements to the SA-NM algorithm include adaptive/optimal factors for shape contraction/expansion/shrinking, an improved shrink method, and shape retension at domain boundaries, greathly improving convergence. For one-dimensional problems, Brent's method is used instead of NM.
  2. An improved version of the Differential Evolution (abbreviated "DE") algorithm, specifically type "DE/rand-to-pbest/1" with archive, by Jingqiao & Arthur (JA-DE). My own improvements to JA-DE include self-adaptive optimal population size reduction and self-adaptive vector length and crossover likelihood. Each search starts with a population of good candidate points selected from a search of the whole multidimensional solution domain using a scrambled low-discrepancy Faure sequence (up to 500 dimensions, and additional dimensions over 500 are calculated as pseudo-random instead of the Faure sequence). This algorithm is slower than NM, but may be more effective at finding global minima. For one-dimensional problems, a low-discrepancy quasi-random search is used instead.
- GlobalMinimize lets you use the above two optimizers in three alternative ways.
  1. Nelder-Mead (NM) only. This is the fastest alternative. This guarantees to find the global minimum for solution domains that are smoothe and are unimodal, but also works very well in solution domains with discontinuous gradient.
  2. Differential-Evolution (DE) followed by Nelder-Mean (NM). This should be your default option, as it increases the likehood of finding the global minimum, and NM 'polishes off' the optimum with a high degree of accuracy.
  3. Indefinite repetitions of DE. This option is a Monte Carlo technique; it increases the chances of finding the global optimum for very 'noisy' solution domains with multiple local minima. You can stop the search at any time by pressing the 'stop' button, after which the NM is automatically run to 'polish off' the solution.
- GlobalMinimize is a normal workbook file, not an Excel AddIn, so it does not require any installation. To use GlobalMinimize, you must have two Excel workbooks open at the same time: the GlobalMinimize workbook and your own workbook file containing the cell value that you wish to minimize.

### Note on using multiple independent variables
- There is no upper limit to the number of "independent variable" cells in GlobalMinimize. It is possible to have hundreds.
- The GlobalMinimize dialogue-box lists four alternative max-min bounds for your independent variables. This is sufficient for most tasks, even if you have hundreds of independent variables. If you experience that 4 alternative max-min bounds is insufficient, please contact me (contact details are below).
- The input fields for min & max bounds must contain a number, and may be positive or negative. You cannot specify a cell reference. The bounds are exclusive (not inclusive), i.e. if you enter min= "0" and max= "1" then the algorithm tries values 0<x<1, not 0≤x≤1. In this case, the values tried will lie between 0.0000⋯001 and  0.9999⋯999.
- The dialogue-box input field between the min & max bounds is for you to specify the "independent variables" with that bound. It has a button that you can click-on to select one or more cells. They can be disjoint cells and/or ranges; select multiple disjoint cells or ranges by clicking on them while pressing the CTRL key.
- GlobalMinimize will change the values in these "independent value" cells in order to minimize the calculated value calculated in the formula in the "objective function" cell. Thus, these "independent variable" cells in your sheet cannot contain formulae, but must instead contain a number.

### Note on using workbooks with multiple sheets or complex problems
- GlobalMinimize remembers the settings (i.e. all the cell addresses in the dialogue-box) so that it can recall them each time you run GlobalMimimize. The settings are saved for each worksheet in the workbook file. Each worksheet has its own GlobalMinimize settings, such that they can be recalled even when you change the sheet name, or duplicate the sheet. A consequence of this is that all the cell references that you specify in the GlobalMinimize dialogue-box must be on the present active sheet (you get a warning if you try to select cells in another sheet).
- Even though the GlobalMinimize dialogue-box permits only cell references on the present active sheet, you can still use GlobalMinimize for complex optimization tasks that encompass multiple sheets (or even multiple workbooks). Simply choose one sheet as your “main” sheet, which you view whenever you run GlobalOptimize, e.g. Sheet1. Then run GlobalOptimize only when you view Sheet1, and in the GlobalOptimize dialogue-box select cells in only on Sheet1. However the formula in the "objective function" cell in Sheet1 can reference cells in the other sheets in your workbook (or other open workbooks), and the cells in the other sheets can reference the “independent variables” cells in Sheet1.
- GlobalMinimize should also work with your own macros. Your objetive function can be a formula with User Defined Function that you program in your own macro.

### Note on discrete vs continuous variables
- GlobalMinimize treats all the independent variables as continuous variables, by default. However, you can cater for discrete variables (ordinate or categorical) simply by using a functions such as INT(), ROUND(), CHOOSE(), etc. in your worksheet This is because GlobalMinimize employs algorithms that work well for functions with a discontinuous or unknown gradient.

### How to automate GlobalMimimize via VBA
- It is now possible to fully automate GlobalMinimize with VBA to solve your own user functions. Contact Peter.Schild@OsloMet.no for details and example-code.

### License and warranty
- Distributed free under the CC BY-ND 4.0 license.
- Provided with no warranty of any kind.

### Author and copyright
- Peter.Schild@OsloMet.no

### References
- Buermen A., Tuma T. "Unconstrained derivative-free optimization by successive approximation". Journal of computational and applied mathematics, vol. 223, pp. 62-74, 2009
- Zhang, Jingqiao, and Arthur C. Sanderson. "JADE: adaptive differential evolution with optional external archive." IEEE Transactions on evolutionary computation 13.5 (2009): 945-958.
- Faure H and Tezuka S (2000) "Another random scrambling of digital (t,s)-sequences Monte Carlo and Quasi-Monte Carlo Methods" Springer-Verlag, Berlin, Germany (eds K T Fang, F J Hickernell and H Niederreiter)

