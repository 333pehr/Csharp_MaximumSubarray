using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace Code
{
    //SEPEHR LATIFI AZAD _ 190254082 _ third-year computer student 
    internal class Program
    {
        static void Main(string[] args)
        {
            // Loading excel files in the C# workspace
            WorkBook Akbank2018 = WorkBook.Load(@"C:\Users\sepeh\OneDrive\Desktop\Algorithm Project\Akbank2018.xlsx");
            WorkBook Akbank2019 = WorkBook.Load(@"C:\Users\sepeh\OneDrive\Desktop\Algorithm Project\Akbank2019.xlsx");
            // Selecting excel sheets for each excel file
            WorkSheet Akbank2018sheet = Akbank2018.GetWorkSheet("isyatirim");
            WorkSheet Akbank2019sheet = Akbank2019.GetWorkSheet("isyatirim");
            // Creating a array for storing data extracted from Excel files
            double[] Akbank2018closedPrise = new double[252];
            double[] Akbank2019closedPrise = new double[250];
            //filling the Arrays created above by using foreach loop
            foreach (var cell in Akbank2018sheet["B2:B252"])
            {
                Akbank2018closedPrise[cell.RowIndex] = Convert.ToDouble(cell.Text);
            }
            foreach (var cell in Akbank2019sheet["B2:B250"])
            {
                Akbank2019closedPrise[cell.RowIndex] = Convert.ToDouble(cell.Text);
            }
            //finding the length of each array
            int n = Akbank2018closedPrise.Length;
            int m = Akbank2019closedPrise.Length;
            //finding the difference or change of each day compared to its next day
            var Akbank2018change = changeFinder(Akbank2018closedPrise);
            var Akbank2019change = changeFinder(Akbank2019closedPrise);
            //results
            Console.WriteLine("Market value of Akbank Finance in the year 2018:(sum of Maximum subArray, start index, end index)");
            Console.WriteLine(findMaxSubarray(Akbank2018change, 1, n - 1));
            Console.WriteLine("\n");
            Console.WriteLine("Market value of Akbank Finance in the year 2019:(sum of Maximum subArray, start index, end index)");
            Console.WriteLine(findMaxSubarray(Akbank2019change, 1, m - 1));
            Console.WriteLine("\n \n");
            Console.WriteLine("****** __2018__ monthly ******");
            Console.WriteLine("January : " + findMaxSubarray(Akbank2018change, 1, 23));
            Console.WriteLine("February : " + findMaxSubarray(Akbank2018change, 24, 42));
            Console.WriteLine("March : " + findMaxSubarray(Akbank2018change, 43, 64));
            Console.WriteLine("April : " + findMaxSubarray(Akbank2018change, 65, 84));
            Console.WriteLine("May : " + findMaxSubarray(Akbank2018change, 85, 106));
            Console.WriteLine("June : " + findMaxSubarray(Akbank2018change, 107, 126));
            Console.WriteLine("July : " + findMaxSubarray(Akbank2018change, 127, 148));
            Console.WriteLine("August : " + findMaxSubarray(Akbank2018change, 149, 166));
            Console.WriteLine("September : " + findMaxSubarray(Akbank2018change, 167, 186));
            Console.WriteLine("October : " + findMaxSubarray(Akbank2018change, 187, 208));
            Console.WriteLine("November : " + findMaxSubarray(Akbank2018change, 209, 230));
            Console.WriteLine("December : " + findMaxSubarray(Akbank2018change, 230, 249));
            Console.WriteLine("\n \n \n");
            Console.WriteLine("****** __2019__ monthly ******");
            Console.WriteLine("January : " + findMaxSubarray(Akbank2019change, 1, 23));
            Console.WriteLine("February : " + findMaxSubarray(Akbank2019change, 24, 42));
            Console.WriteLine("March : " + findMaxSubarray(Akbank2019change, 43, 64));
            Console.WriteLine("April : " + findMaxSubarray(Akbank2019change, 65, 84));
            Console.WriteLine("May : " + findMaxSubarray(Akbank2019change, 85, 106));
            Console.WriteLine("June : " + findMaxSubarray(Akbank2019change, 107, 126));
            Console.WriteLine("July : " + findMaxSubarray(Akbank2019change, 127, 148));
            Console.WriteLine("August : " + findMaxSubarray(Akbank2019change, 149, 166));
            Console.WriteLine("September : " + findMaxSubarray(Akbank2019change, 167, 186));
            Console.WriteLine("October : " + findMaxSubarray(Akbank2019change, 187, 208));
            Console.WriteLine("November : " + findMaxSubarray(Akbank2019change, 209, 230));
            Console.WriteLine("December : " + findMaxSubarray(Akbank2019change, 230, 249));

            Console.WriteLine("****** __2018__ Weekly January ******");
            Console.WriteLine("1'st week : " + findMaxSubarray(Akbank2018change, 1, 4));
            Console.WriteLine("2'nd week : " + findMaxSubarray(Akbank2018change, 5, 9));
            Console.WriteLine("3'rd week : " + findMaxSubarray(Akbank2018change, 10, 14));
            Console.WriteLine("4'th week : " + findMaxSubarray(Akbank2018change, 15, 22));
            Console.WriteLine("\n");
            Console.WriteLine("****** __2018__ Weekly September ******");
            Console.WriteLine("1'st week : " + findMaxSubarray(Akbank2018change, 168, 171));
            Console.WriteLine("2'nd week : " + findMaxSubarray(Akbank2018change, 172, 176));
            Console.WriteLine("3'rd week : " + findMaxSubarray(Akbank2018change, 177, 181));
            Console.WriteLine("4'th week : " + findMaxSubarray(Akbank2018change, 182, 186));
            Console.WriteLine("\n");
            Console.WriteLine("****** __2019__ Weekly January ******");
            Console.WriteLine("1'st week : " + findMaxSubarray(Akbank2019change, 1, 3));
            Console.WriteLine("2'nd week : " + findMaxSubarray(Akbank2019change, 4, 8));
            Console.WriteLine("3'rd week : " + findMaxSubarray(Akbank2019change, 9, 13));
            Console.WriteLine("4'th week : " + findMaxSubarray(Akbank2019change, 14, 22));
            Console.WriteLine("\n");
            Console.WriteLine("****** __2019__ Weekly September ******");
            Console.WriteLine("1'st week : " + findMaxSubarray(Akbank2019change, 165, 169));
            Console.WriteLine("2'nd week : " + findMaxSubarray(Akbank2019change, 170, 174));
            Console.WriteLine("3'rd week : " + findMaxSubarray(Akbank2019change, 175, 179));
            Console.WriteLine("4'th week : " + findMaxSubarray(Akbank2019change, 180, 184));

            Console.ReadKey();
        }

        //function to find difference or change of each day compared to its next day
        static double[] changeFinder(double[] array)
        {
            //creating a temporary array to store data
            double[] result = new double[array.Length];
            for(int i=0; i<array.Length -1; i++)
            {
                //finding difference by subtract operator
                result[i] = array[i] - array[i+1];
                result[i] = (result[i] * -1);
            }
            return result;
        }

        // Function to find the maximum subarray and its interval using divide and conquer
        static (double sum, int start, int end) findMaxSubarray(double[] arr, int low, int high)
        {
            // If there is only one element in the array, return it
            if (low == high)
                return (arr[low], low, high);

            // Find the midpoint of the array
            int midindex = (low + high) / 2;

            // Find the maximum subarray and its interval in the left subarray
            var left = findMaxSubarray(arr, low, midindex);

            // Find the maximum subarray and its interval in the right subarray
            var right = findMaxSubarray(arr, midindex + 1, high);

            // Find the maximum subarray and its interval that includes the midpoint
            var mid = findMaxCrossingSubarray(arr, low, midindex, high);

            // Return the maximum subarray and its interval of the three values
            if (left.sum >= right.sum && left.sum >= mid.sum)
                return left;
            else if (right.sum >= left.sum && right.sum >= mid.sum)
                return right;
            else
                return mid;
        }

        // Function to find the maximum subarray and its interval that includes the midpoint
        static (double sum, int start, int end) findMaxCrossingSubarray(double[] arr, int low, int mid, int high)
        {
            // Initialize the left and right sums to the minimum possible value
            double leftSum = double.MinValue;
            double rightSum = double.MinValue;

            // Initialize the left and right indices to the current midpoint
            int leftIndex = mid;
            int rightIndex = mid + 1;

            // Find the maximum subarray sum that starts from the midpoint and extends to the left
            double sum = 0;
            for (int i = mid; i >= low; i--)
            {
                sum += arr[i];
                if (sum > leftSum)
                {
                    leftSum = sum;
                    leftIndex = i;
                }
            }

            // Find the maximum subarray sum that starts from the midpoint and extends to the right
            sum = 0;
            for (int i = mid + 1; i <= high; i++)
            {
                sum += arr[i];
                if (sum > rightSum)
                {
                    rightSum = sum;
                    rightIndex = i;
                }
            }

            // Return the maximum subarray and its interval
            return (leftSum + rightSum, leftIndex, rightIndex);
        }

    }
}
