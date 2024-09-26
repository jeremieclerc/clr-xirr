using System;
using System.Collections;
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
        guess = 0.1;                      // Default guess value
    }

    public void Accumulate(SqlDouble values, SqlDateTime dates, SqlDouble guess)
    {
        // Check if the date or the amount are null or if amount=0. In any of these cases, ignore this row
        if ((!dates.IsNull) && (!values.IsNull) && (!values.Equals((SqlDouble)0)))
        {
            int days = ((DateTime)dates - DateTime.Today).Days;

            // Check if that date is already in the list
            if (irrElements.IndexOfKey(days) == -1)
                irrElements.Add(days, Math.Truncate(values.Value * 10000) / 10000);
            else
            {
                // if it is, then add the amount
                Double originalAmount = (Double)(irrElements[days]);
                irrElements[days] = originalAmount + (Math.Truncate(values.Value * 10000) / 10000);
            }
        }

        if (!guess.IsNull)
            this.guess = guess.Value;
    }

    public void Merge(XIRR Group)
    {
        guess = Group.guess;

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
            xirr = XirrCalculator.calculateIRR(values, days, guess);

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
        public static double calculateIRR(IList values, IList days, double guess)
        {
            double rate = guess;
            double tol = 0.000001; // Tolerance level for convergence
            int maxIterations = 100; // Maximum number of iterations allowed
            double npv = 0.0;
            double npvDerivative = 0.0;
            double minRate = -0.9999999; // Minimum allowable rate (rate cannot be less than -1)
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
                        return double.NaN;
                    }

                    double denom = Math.Pow(ratePlusOne, t);

                    if (denom == 0.0)
                    {
                        // Avoid division by zero
                        denom = 1e-10;
                    }

                    npv += (double)values[i] / denom;
                    npvDerivative -= t * (double)values[i] / (denom * ratePlusOne);
                }

                if (Math.Abs(npv) < tol)
                {
                    // NPV is close enough to zero; convergence achieved
                    return rate;
                }

                if (npvDerivative == 0.0)
                {
                    // Derivative is zero; cannot proceed
                    return double.NaN;
                }

                double newRate = rate - npv / npvDerivative;

                if (double.IsNaN(newRate) || double.IsInfinity(newRate))
                {
                    // Computation failed; return NaN
                    return double.NaN;
                }

                // Ensure the rate stays above the minimum allowable rate
                if (newRate <= minRate)
                {
                    newRate = (rate + minRate) / 2.0;
                }

                rate = newRate;
                iteration++;
            }

            // Maximum iterations exceeded without convergence
            return double.NaN;
        }

    }
}
