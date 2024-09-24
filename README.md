
# XIRR Calculation using Newton-Raphson Method

## Overview

This project implements an SQL User Defined Aggregate (UDA) function in C# to calculate the Internal Rate of Return (IRR) using the Newton-Raphson method. The `XIRR` function is designed to handle irregular cash flow intervals and approximates the IRR for a given set of cash flows and dates. The objective was to be as close as possible to the XIRR function of Microsoft Excel. [See Microsoft documentation.](https://support.microsoft.com/en-us/office-xirr-function-de1242ec-6477-445b-b11b-a303ad9adc9d)

## Getting Started

You need to have :

- Visual Studio Professional (2016-2022)
- SSDT (SQL Server Data Tools)
- SDK .NET version 7.0 or above (.Net desktop development)
- SQL Server 2014, 2016, 2017, 2019, 2022 or Managed Instance

## Quick install

Execute the file: `AssemblyXirr.sql` located inside the [Releases tab.](https://github.com/jeremieclerc/clr-xirr/releases)
This will create the assembly and the function `XIRR` associated.

## Full installation

- Clone the repo and open the solution with Visual Studio
- Build the assembly
- Load the assembly into SQL Server

## Usage

This function is strongly inspired by Excel's XIRR function.
It accepts 3 parameters:

- Amounts
- Dates
- Guess

## Example

```sql
SELECT dbo.XIRR(AmountColumn, DateColumn, 0.1)
FROM FinancialTable
```

## Code Structure

### 1. `XIRR` Struct

The main structure, `XIRR`, contains the following methods:

- **`Init()`**: Initializes the IRR elements and sets the initial guess value.
- **`Accumulate(SqlDouble values, SqlDateTime dates, SqlDouble guess)`**: Accumulates the dates and amounts into a sorted list and sets the initial guess value.
- **`Merge(XIRR Group)`**: Merges another `XIRR` group into the current instance, combining the IRR elements and updating the guess value.
- **`Terminate()`**: Calculates the final IRR using the `calculateIRR` method of the `XirrCalculator` class and returns the result as `SqlDouble`.

### 2. `XirrCalculator` Class

This class contains the logic for calculating the IRR using the Newton-Raphson method.

`calculateIRR(IList values, IList days, double guess)`

This method performs the IRR calculation with the following steps:

1. Initializes the IRR rate with the provided guess value.
2. Iterates up to a maximum of 100 times or until the NPV is within a tolerance of `1e-7`.
3. Calculates both the NPV and its derivative at the current rate.
4. Updates the rate using the Newton-Raphson formula:

```math
   \text{{newRate}} = \text{{rate}} - \frac{{\text{{NPV}}}}{{\text{{NPV'}}}}
```

### 3. Serialization Methods

- **`Read(BinaryReader r)`**: Deserializes the state of the `XIRR` structure from a binary stream.
- **`Write(BinaryWriter w)`**: Serializes the state of the `XIRR` structure to a binary stream.

## Contributing

Feel free to fork this repository and make your own contributions. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT License](LICENSE)

## Authors

- [Arthur Meyniel](https://github.com/ArthurMynl)
- [Jérémie Clerc](https://github.com/jeremieclerc)
