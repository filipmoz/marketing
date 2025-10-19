# Core Statistical Functions of Marketing

## Descriptive Statistics
- **AVERAGE** - Calculate mean scores
  ```
  =AVERAGE('Survey Data'!B:B)
  ```
- **STDEV** - Standard deviation
  ```
  =STDEV('Survey Data'!B:B)
  ```
- **MIN/MAX** - Minimum and maximum values
  ```
  =MIN('Survey Data'!B:B)  |  =MAX('Survey Data'!B:B)
  ```

## Counting & Filtering
- **COUNTIF** - Count matching criteria
  ```
  =COUNTIF('Survey Data'!N:N,"Female")
  ```
- **COUNTIFS** - Count with multiple criteria
  ```
  =COUNTIFS('Survey Data'!P:P,"18 to 34",'Survey Data'!I:I,">=5")
  ```

## Conditional Averages
- **AVERAGEIF** - Average with one condition
  ```
  =AVERAGEIF(A:A,"Married",B:B)
  ```
- **AVERAGEIFS** - Average with multiple conditions
  ```
  =AVERAGEIFS(B:B,A:A,"Female",C:C,"18 to 34")
  ```

## Hypothesis Testing
- **CHITEST** - Chi-square test (categorical associations)
- **T.TEST** - T-test (comparing two group means)
- **ANOVA** - Comparing three or more groups

## Analysis Tools
- **Crosstabs** - Age Group × Innovator Personality (High/Low)
- **Charts** - Pie charts (gender), bar charts (age), line charts (attitudes)
