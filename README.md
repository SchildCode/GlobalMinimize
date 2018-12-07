# GlobalMinimize
### Features
- A Microsoft Excel macro with state-of-the art optimization algorithms that outperform the Excel Solver AddIn when the solution domain contains multiple local minima or regions with constant cost (no gradient). 
- The application contains improved versions of the following two optimizers: 
-- (i) A self-adaptive version of Downhill Simplex algorithm by Nelder-Mead (abbreviated "NM"). This algorithm is faster than DE, and uses successive-approximation to guarantee convergence at minima. For one-dimensional problems the Brent's method is used instead.
-- (ii) An improved version of the Differential Evolution (abbreviated "DE") algorithm, specifically DE/rand-to-pbest/1 JADE with archive. My improvements include self-adaptive optimal population size reduction and self-adaptive vector length and crossover likelihood. This algorithm is very effective at finding global minimuma. Each search starts with a polulation of good candidate points selected from a search of the whole multidimensional solution domain using a scrambled low-discrepancy Faure sequence.
- GlobalMinimize lets you use the above two optimizers in three alternative ways.
  (a) Nelder-Mead (NM) only. This is the fastest alternative. This guarantees to find the global minimum for solution domains that are smoothe and are unimodal, but also works very well in solution domains with discontinuous gradient.
  (b) Differential-Evolution (DE) followed by Nelder-Mean (NM). This should be your default option, as it increases the likehood of finding the global minimum, and NM 'polishes off' the optimum with a high degree of accuracy.
  (c) Indefinite repetitions of DE. This option is a Monte Carlo technique; it increases the chances of finding the global optimum for very 'noisy' solution domains with multiple local minima. You can stop the search at any time by pressing the 'stop' button, after which the NM is automatically run to 'polish off' the solution.

### License and warranty
- Distributed free under the CC RY-ND 4.0 license.
- Provided with no warranty of any kind.

### Author and copyright
- Peter.Schild@OsloMet.no

### References
- Buermen A., Tuma T. "Unconstrained derivative-free optimization by successive approximation". Journal of computational and applied mathematics, vol. 223, pp. 62-74, 2009
- Faure H and Tezuka S (2000) "Another random scrambling of digital (t,s)-sequences Monte Carlo and Quasi-Monte Carlo Methods" Springer-Verlag, Berlin, Germany (eds K T Fang, F J Hickernell and H Niederreiter)
- Zhang, Jingqiao, and Arthur C. Sanderson. "JADE: adaptive differential evolution with optional external archive." IEEE Transactions on evolutionary computation 13.5 (2009): 945-958.
