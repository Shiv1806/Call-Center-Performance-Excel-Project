# Project Details

## Objective

Build an Excel-based call center performance reporting solution that converts detailed call records into usable operational insights.

## Dataset

The workbook contains 1,000 call records and customer information. The data includes operational measures such as duration, purchase amount, satisfaction rating and representative, alongside customer attributes including city, gender and age.

## Analytical Layers

### Data Layer
The `Data` sheet contains the working call-center dataset and calculated fields.

### Analysis Layer
The `Pivot Tables` sheet contains multiple PivotTables that answer different operational questions.

Examples include monthly trends, weekday activity, representative performance, rating distributions, city/gender patterns, and customer/representative amount analysis.

### Presentation Layer
The `Dashboard Customer Care Center` sheet consolidates the analysis into a visual dashboard with KPI cards, charts and slicer-driven interaction.

### Supporting Layer
The `Assets` sheet contains representative image / lookup resources used for presentation.

## KPI Framework

The dashboard uses five headline measures:

- Total Calls
- Total Amount
- Total Duration
- Average Rating
- Happy Customers

## Representative Analysis

The representative analysis includes both call volume and purchase amount. The workbook also contains supporting calculations for:

- Selected representative
- Percentage of calls
- Call rank
- Amount rank
- Representative-specific lookup values

## Visualization

The dashboard uses multiple chart types to show:

- Monthly trend
- Weekday trend
- Gender comparison
- Rating distribution
- Representative calls
- Representative amount

The design separates the calculation layer from the presentation layer, making the workbook easier to navigate and maintain.

## Interactivity

The workbook contains a Representative slicer connected to multiple PivotTable views. This allows the user to change the representative selection and examine the corresponding dashboard calculations.

## Portfolio Value

This project is suitable as a portfolio example because it demonstrates a complete workflow:

1. Prepare data.
2. Create analytical fields.
3. Aggregate with PivotTables.
4. Build supporting calculations.
5. Visualize the results.
6. Add interactive filtering.
7. Present the result as a dashboard.
