using Microsoft.SqlServer.Server;
using System;
using System.Collections;
using System.Data.SqlTypes;

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.Native)]
public struct XIRR
{
    public SortedList irrElements;
    public Double guess;
    public void Init()
    {
        irrElements = new SortedList();   // dates/values list
        guess = 12345;                    // Random value to approximate IRR
    }

    public void Accumulate(SqlDateTime dates, SqlDouble amounts, SqlDouble guess)
    {
        // Check if the date or the amount are null or if amount=0. In any of these cases, ignore this row
        if ((!dates.IsNull) && (!amounts.IsNull) && (!amounts.Equals((SqlDouble)0)))
        {
            // Get the distance in days between the date and today
            int days = ((DateTime)dates - DateTime.Today).Days;

            // Check if that date is already in the list
            if (irrElements.IndexOfKey(days) == -1)
                // if it is not, add it
                irrElements.Add(days, (Double)amounts);
            else
            {
                // if it is, then add the amount
                Double originalAmount = (Double)(irrElements[days]);
                irrElements[days] = originalAmount + (Double)amounts;
            }
        }

        // Initialize the approximate value
        if (this.guess == 12345)
            this.guess = (Double)guess.Value;
    }

    public void Merge (XIRR Group)
    {
        // Put the merge code here
        // Merge the IRR approximate
        guess = Group.guess;

        // For each element of the Group SortedList, append or update the local list
        SortedList groupList = Group.irrElements;
        IList groupKeyList = groupList.GetKeyList();
        IList groupValueList = groupList.GetValueList();

        for (int i = 0; i < groupList.Count; i++)
        {
            // Check if that date is already in the local list
            if (irrElements.IndexOfKey(groupKeyList[i]) == -1)
                // if it is not, add it
                irrElements.Add(groupKeyList[i], (Double)groupValueList[i]);
            else
            {
                // if it is, then add the amount
                Double originalAmount = (Double)(irrElements[groupKeyList[i]]);
                irrElements[groupKeyList[i]] = originalAmount + (Double)groupValueList[i];
            }
        }
    }

    public SqlDouble Terminate ()
    {
        // Gets the list of keys and the list of values.
        IList days = this.irrElements.GetKeyList();
        IList payments = this.irrElements.GetValueList();

        bool noPositiveValueExists = true;
        bool noNegativeValueExists = true;
        int i = 0;

        // Loop while we have not found a positive and a negative value in all the list
        while ((i < payments.Count) && (noPositiveValueExists || noNegativeValueExists))
        {
            if ((double)payments[i] > 0)
                noPositiveValueExists = false;

            if ((double)payments[i] < 0)
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
            xirr = XirrCalculator.calculateIRR(guess, payments, days);

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

    // This is a place-holder member field
    // Custom serialization read method
    public void Read(System.IO.BinaryReader r)
    {

        // Read the two parameters given to the function
        guess = r.ReadDouble();

        // Read the list of dates/amounts

        // Create a new sorted list
        this.irrElements = new SortedList();

        // Get the number of values that were serialized
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

        // Store the two parameters given to the function
        w.Write(guess);

        // Store the list of dates/amounts
        // Write the number of values in the list
        w.Write(this.irrElements.Count);

        // Gets the list of keys and the list of values.
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
        public static double calculateIRR(double guess, IList payments, IList days)
        {
            // This function calculates the Internal Rate of Return (IRR) for the given data.
            double irr = 0;
            double tolerance = 0.00001; // Tolerance level for the convergence of the IRR calculation.
            double rate = guess; 
            double lastRate; // Variable to store the IRR value from the previous iteration.
            double error = 100.0; // Initial error set to a high value to start the loop.
            double maximumError = 1E30; // Added to check if the error diverges beyond this value, indicating failure.
            double rateStep = 1.0; 
            double residual = 0.99; 
            double lastResidual = 1.0;
            int maxIterations = 100;

            int i = 0; // Iteration counter.
            lastRate = rate; // Initialize the previous rate with the initial guess.

            while ((error > tolerance) && (i < maxIterations)) // Loop until the error is within tolerance or max iterations are reached.
            {
                // Store the result from the previous iteration.
                lastResidual = residual;
                lastRate = rate;

                // Calculate the net present value of the payments using the current rate.
                residual = calculateReturn(rate, payments, days);

                // Calculate the absolute difference between the last and current residuals (error).
                error = Math.Abs((lastResidual - residual));

                // Prepare for the next iteration.
                if (residual >= 0)
                    rate += rateStep;
                else 
                {
                    rateStep /= 2;
                    rate -= rateStep;
                }

                i++;

                // If the error diverges and exceeds the maximum error threshold, return NaN (not a number).
                if (error > maximumError)
                    return double.NaN;
            }

            irr = lastRate; // Store the last calculated rate as the final IRR.

            // If the calculated IRR is positive or negative infinity, return NaN.
            if (irr.Equals(Double.PositiveInfinity) || irr.Equals(Double.NegativeInfinity))
                return double.NaN;

            // If the function has not converged within the allowed iterations, return NaN.
            if (i == maxIterations)
                return double.NaN;

            return irr;
        }


        private static double calculateReturn(double rate, IList payments, IList days)
        {
            // This function calculates the compounded interest based on the list of dates/payments at a given rate.
            double returnValue = 0;

            for (int i = 0; i < payments.Count; i++) // Iterate over all payments in the list.
            {
                // Check if the calculated power value is zero, which can cause a division error.
                if (Math.Pow((double)(1.0 + rate), (double)(((Int32)days[i] - (Int32)days[0]) / 365.0)).Equals(0.0))
                    // If the calculated power is zero, divide the payment by a small number (0.1) to avoid division by zero.
                    returnValue += (double)payments[i] / 0.1;
                else
                    // Otherwise, calculate the present value of the payment using the formula:
                    // payment / (1 + rate) ^ ((days[i] - days[0]) / 365)
                    returnValue += (double)payments[i] / Math.Pow((double)(1.0 + rate), (double)(((Int32)days[i] - (Int32)days[0]) / 365.0));
            }
            return returnValue;
        }

    }
}
