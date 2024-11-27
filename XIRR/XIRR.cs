using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;

[Serializable]
[SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize = -1)]
public struct XIRR : IBinarySerialize
{
    public SortedList irrElements;
    public Double guess;

    public void Init()
    {
        irrElements = new SortedList();   // dates/values list
    }

    public void Accumulate(SqlDouble values, SqlDateTime dates)
    {
        // Check if the date or the amount are null or if amount=0. In any of these cases, ignore this row
        if ((!dates.IsNull) && (!values.IsNull) && (!values.Equals((SqlDouble)0)))
        {
            int days = ((DateTime)dates - DateTime.Today).Days;

            // Check if that date is already in the list
            if (irrElements.IndexOfKey(days) == -1)
                irrElements.Add(days, values.Value);
            else
            {
                // if it is, then add the amount
                Double originalAmount = (Double)(irrElements[days]);
                irrElements[days] = originalAmount + values.Value;
            }
        }
    }

    public void Merge(XIRR Group)
    {
        SortedList groupList = Group.irrElements;
        IList groupKeyList = groupList.GetKeyList();
        IList groupValueList = groupList.GetValueList();

        for (int i = 0; i < groupList.Count; i++)
        {
            if (irrElements.IndexOfKey(groupKeyList[i]) == -1)
                irrElements.Add(groupKeyList[i], (Double)groupValueList[i]);
            else
            {
                Double originalAmount = (Double)(irrElements[groupKeyList[i]]);
                irrElements[groupKeyList[i]] = originalAmount + (Double)groupValueList[i];
            }
        }
    }

    public SqlDouble Terminate()
    {
        IList days = this.irrElements.GetKeyList();
        IList values = this.irrElements.GetValueList();

        bool noPositiveValueExists = true;
        bool noNegativeValueExists = true;
        int i = 0;

        // Loop while we have not found a positive and a negative value in all the list
        while ((i < values.Count) && (noPositiveValueExists || noNegativeValueExists))
        {
            if ((double)values[i] > 0)
                noPositiveValueExists = false;

            if ((double)values[i] < 0)
                noNegativeValueExists = false;

            i++;
        }

        double xirr;

        // If no positive or negative value exists, return NULL, otherwise calculate the IRR
        if (noPositiveValueExists || noNegativeValueExists)
        {
            return SqlDouble.Null;
        }
        else
        {
            guess = XirrCalculator.AutoGuess(values);
            xirr = XirrCalculator.CalculateIRR(values, days, guess);

            // Check if xirr indicates an error (e.g., calculation did not converge)
            if (double.IsNaN(xirr))
            {
                return SqlDouble.Null;
            }
            else
            {
                return new SqlDouble(xirr);
            }
        }
    }

    // Custom serialization read method
    public void Read(System.IO.BinaryReader r)
    {
        guess = r.ReadDouble();

        this.irrElements = new SortedList();

        int j = r.ReadInt32();

        // Loop and add each serialized value to the list
        for (int i = 0; i < j; i++)
        {
            this.irrElements.Add(r.ReadInt32(), r.ReadDouble());
        }
    }

    // Custom serialization write method
    public void Write(System.IO.BinaryWriter w)
    {
        w.Write(guess);
        w.Write(this.irrElements.Count);

        IList myKeyList = this.irrElements.GetKeyList();
        IList myValueList = this.irrElements.GetValueList();

        // Prints the keys in the first column and the values in the second column.
        for (int i = 0; i < this.irrElements.Count; i++)
        {
            w.Write((Int32)myKeyList[i]);
            w.Write((Double)myValueList[i]);
        }
    }

    public static class XirrCalculator
    {

        public static double AutoGuess(IList values)
        {
            double sumCashflows = 0;
            for (int i = 0; i < values.Count; i++)
            {
                sumCashflows += (double)values[i];
            }
            if (sumCashflows >= 0)
                return 0.1;
            else
                return -0.1;
        }

        public static double CalculateIRR(IList values, IList days, double guess)
        {
            double rate = guess;
            double tol = 1e-6; // Tolerance for convergence
            int maxIterations = 100; // Maximum number of iterations

            // Attempt to find IRR using Newton-Raphson method first
            rate = CalculateIRRUsingNewtonRaphson(values, days, guess, tol, maxIterations);
            if (!double.IsNaN(rate) && rate > -1)
                return rate;

            rate = CalculateIRRUsingLegacy(values, days, guess);
            if (!double.IsNaN(rate) && rate > -1)
                return rate;

            // If all methods fail, return NaN as a fallback
            return double.NaN;
        }

        private static double CalculateIRRUsingNewtonRaphson(IList values, IList days, double guess, double tol, int maxIterations)
        {
            double rate = guess;
            double npv, npvDerivative;
            int iteration = 0;

            while (iteration < maxIterations)
            {
                npv = 0.0;
                npvDerivative = 0.0;

                for (int i = 0; i < values.Count; i++)
                {
                    double t = ((Int32)days[i] - (Int32)days[0]) / 365.0;
                    double ratePlusOne = rate + 1.0;

                    if (ratePlusOne <= 0.0)
                    {
                        // Cannot compute with rate + 1 <= 0
                        ratePlusOne = 1e-10;
                    }

                    double denom = Math.Pow(ratePlusOne, t);
                    if (denom == 0.0)
                    {
                        denom = 1e-10; // Avoid division by zero
                    }

                    npv += (double)values[i] / denom;
                    npvDerivative -= t * (double)values[i] / (denom * ratePlusOne);
                }

                if (Math.Abs(npv) < tol)
                {
                    return rate; // Convergence achieved
                }

                if (Math.Abs(npvDerivative) < tol)
                {
                    if (npvDerivative < 0)
                        npvDerivative = -0.1;
                    else
                        npvDerivative = 0.1;
                }

                rate = rate - npv / npvDerivative; // Newton-Raphson formula

                iteration++;
            }

            return double.NaN; // Max iterations exceeded
        }

        public static double CalculateIRRUsingLegacy(IList values, IList days, double guess)
        {
            // This function calculates the Internal Rate of Return (IRR) for the given data.
            double irr = 0;
            double tolerance = 1e-6; // Tolerance level for the convergence of the IRR calculation.
            double rate = guess;
            double lastRate; // Variable to store the IRR value from the previous iteration.
            double error = 100.0; // Initial error set to a high value to start the loop.
            double rateStep = 1.0;
            double npv = 0.99;
            double lastNpv = 1.0;
            int maxIterations = 100;
            int iteration = 0; // Iteration counter.
            lastRate = rate; // Initialize the previous rate with the initial guess.

            while ((error > tolerance) && (iteration < maxIterations)) // Loop until the error is within tolerance or max iterations are reached.
            {
                // Store the result from the previous iteration.
                lastNpv = npv;
                lastRate = rate;

                // Calculate the net present value (NPV) of the payments using the current rate.
                double returnValue = 0.0;
                for (int j = 0; j < values.Count; j++) // Iterate over all payments in the list.
                {
                    // Calculate the exponent based on the difference in days.
                    double exponent = ((int)days[j] - (int)days[0]) / 365.0;

                    // Calculate the denominator (1 + rate) raised to the exponent.
                    double denominator = Math.Pow(1.0 + rate, exponent);

                    // Check if the denominator is zero to avoid division by zero.
                    if (denominator.Equals(0.0))
                    {
                        // If the denominator is zero, use a small number to prevent division by zero.
                        returnValue += (double)values[j] / 0.1;
                    }
                    else
                    {
                        // Otherwise, calculate the present value of the payment.
                        returnValue += (double)values[j] / denominator;
                    }
                }

                npv = returnValue;

                // Update the rate based on the NPV.
                if (npv >= 0)
                {
                    rate += rateStep;
                }
                else
                {
                    rateStep /= 2;
                    rate -= rateStep;
                }

                // Update the error (optional: you might want to define how error is calculated).
                // For example, you could use the absolute difference between current and last NPV.
                error = Math.Abs(npv - lastNpv);

                iteration++; // Increment the iteration counter.
            }

            irr = lastRate; // Store the last calculated rate as the final IRR.

            // If the calculated IRR is positive or negative infinity, return NaN.
            if (irr.Equals(double.PositiveInfinity) || irr.Equals(double.NegativeInfinity))
                return double.NaN;

            // If the function has not converged within the allowed iterations, return NaN.
            if (iteration == maxIterations)
                return double.NaN;

            return irr;
        }

    }
}
