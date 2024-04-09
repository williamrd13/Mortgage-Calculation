### Statement of purpose

Imagine I'm working in a financial institution and my goal is to give my customers a basic overview of real-estate or car loans. Given a certain expected loan amount, perhaps loan length, annual interest rate, and start date together, I would like to tell them what the calculated monthly payment would be immediately. Additionally, a complete schedule of payments would also be appreciated. There are plenty of online services which can achieve that, but unfortunately, those are not ads-free. So my goal here is to make a simple tool to calculate and display the monthly payment, and export a payment schedule to an excel file.

### Formula

There are proofs and mortgage formulae online. I will still derive one for my own use.

In my proof I will use the definite geometric series. Note that |x| < 1

$$
\begin{aligned}
f(x)&=1+x+x^2+...+x^\{(n-1)}\\
&=(1+x+x^2+...+x^\{(n-1)})*\frac{(x-1)}{(x-1)}\\
&=\frac{(x+x^2+x^3+...+x^n)-(1+x+x^2+...+x^\{(n-1)})}{x-1}\\
&=\frac{x^n-1}{x-1}\\
&=\frac{1-x^n}{1-x}\\
\end{aligned}
$$

Now treat x as the constant monthly payment. The loan amount and each monthly payment will roll with interests according to their distance to the maturity date.
For simplicity, I derive ANNUAL payments here instead. Let's begin with future value (FV) of each cash flow on the maturity date first.
So x = annual payment, r = annual interest rate, n = number of years

$$
\begin{aligned}
\{FV\ of\ loan\ amount} = \{loan\ amount*(1+r)^\{n}}\\
\{FV\ of\ all\ payments} = x*(1+r)^\{n}+...+x*(1+r)^2+x*(1+r)^1\\
\{PV\ of\ loan\ amount} = \{loan\ amount}\\
\{loan\ amount} = {PV\ of\ all\ payments} &= \frac{x}{1+r}+\frac{x}{(1+r)^2}+...+\frac{x}{(1+r)^\{n}}\\
&=(\frac{x}{1+r})+(\frac{x}{1+r})\*(\frac{1}{1+r})+...+(\frac{x}{1+r})\*(\frac{1}{(1+r)^\{(n-1)}})\\
&=(\frac{x}{1+r})\*\frac{1-\frac{1}{(1+r)^n}}{1-\frac{1}{1+r}}\\
&=(\frac{x}{1+r})\*(1-\frac{1}{(1+r)^n})\*\frac{1}{\frac{1+r-1}{1+r}}\\
&=(\frac{x}{1+r})\*(1-\frac{1}{(1+r)^n})\*\frac{1+r}{r}\\
&=(\frac{x}{r})\*(1-\frac{1}{(1+r)^n})\\
\end{aligned}
$$

The convention for converting annual payments into monthly payments is that the interest rate for each period becomes 1/12, and it rolls over each month (n\*12) periods. 
For example, for a 30-year loan with anuual rate of 4.8%, we use r = 4.8%/12 = 0.4% and n = 30\*12 = 360 to calculate x, where x becomes the monthly payment.

### Building with Jupyter

There would be multiple steps to do. First, I need to make the calculation work. This won't be difficult after proving the formula. After calculating the monthly payment, I would like to show the balance, interest payment, and balance deduction as well so my customer would have a better view of their payment schedule (better with dates). All of these can also be calculated iteratively. Then I need to setup a simple user interface which a user (like me, my colleagues, or even a random person) can input information and get the result using simple clicks.

When exporting the calculated data to excel, I find that they don't look very organized in an excel file. That's why I need to incorporate openpyxl package to manipulate the overall format and individual cell format. The excel file should be ready to read as soon as it is opened.

The final task to do would be bug-fixing. There are certain potential errors may occur when using the app. These errors may cause unexpected crash which could be annoying. There are several following scenarios which would cause errors.
- Empty information (lack of necessary element to calculate)
- Invalid input
- Excel file already opened (permission error)

To solve this, I just need to put steps in try blocks and test if the error message shows correctly.

### To the next level

What if a customer demands a wierd plan which the monthly payment increases by a constant each month? This sounds ridiculus and it's not commonly seen in a real-world scenario. We can still solve this problem with a augmented proof and adjust our app.

### New Formula

[Formula proof]

### Building with Jupyter

The important thing to change is the monthly payment because it is no longer a constant. As shown in the proof, it is the combination of a constant payment plus an increasing payment. To do this, we need to add the monthly payment into our iteration as well.

### Summary
