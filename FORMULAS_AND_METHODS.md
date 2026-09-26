# Formulas and Methods

## Derived Columns

### Financial Year

```excel
=IF(MONTH([@[Date of Call]])<=6,YEAR([@[Date of Call]]),YEAR([@[Date of Call]])+1)
```

### Day of Week

```excel
=TEXT([@[Date of Call]],"DDDD")
```

### Duration Bucket

```excel
=IFS(
[@[Duration]]<=10,"Under 10 mins",
[@[Duration]]<=30,"10 to 30 mins",
[@[Duration]]<=60,"30 to 60 mins",
[@[Duration]]<=120,"1 to 2 hours",
TRUE,"More than 2 hours"
)
```

### Rounded Rating

```excel
=ROUND([@[Satisfaction Rating]],0)
```

## Lookup Examples

Selected representative call count:

```excel
=XLOOKUP(C54,H22:H26,I22:I26)
```

Selected representative amount:

```excel
=XLOOKUP(C54,H22:H26,J22:J26)
```

## Ranking

Call rank:

```excel
=RANK.AVG(K33,N33:N37)
```

Amount rank:

```excel
=RANK.AVG(K34,O33:O37)
```

## Excel Techniques Used During the Project

### Top 5

```excel
=TAKE(SORT(A2:F261,6,-1),5)
```

### Bottom 5

```excel
=TAKE(SORT(A2:F261,6,1),5)
```

### Filtered result

```excel
=IF(E34=$C$54,F34,NA())---- This will show the representative image, call and amount data only if single slicer button is selected,
                            else no result will be displayed.
```

### Conditional Highlighting

```excel
=N$3='Pivot Tables'!$E$67--- This will highlight a column in the pivot table according to the selected
                             representative on the slicer.
```

### Customise the size of data bars.

```excel
=MAX('Dashboard Customer Care Center'!$N$5:$R$21)*2---- scale down the size of data bars according to maximum value in the range. Makes the data more readable.
```

## Design Method

The workbook separates:

**Data → PivotTables → Helper Calculations → Dashboard**

This keeps the dashboard presentation layer independent from the detailed calculation layer.
