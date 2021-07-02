package exceptionHandling;

import java.util.InputMismatchException;
import java.util.Scanner;

import static java.lang.Math.abs;
import static java.lang.Math.sqrt;

public class QuadraticEquation {

    public static void main(String[] args) {
        boolean condition = true;
        int firstNumber = 0;
        double secondNumbner = 0.0;
        double thirtNumber = 0.0;

        while (condition == true) {
            try {
                Scanner scanner = new Scanner(System.in);
                System.out.println("Enter first number : ");
                firstNumber = scanner.nextInt();
                if (firstNumber == 0) {
                    throw new InputMismatchException();
                }
                System.out.println("Enter second number : ");
                secondNumbner = scanner.nextInt();
                System.out.println("Enter third number : ");
                thirtNumber = scanner.nextInt();
                condition = false;

            } catch (InputMismatchException exception) {
                System.out.println("you are putting wrong number format, Please put valid Number");
            }
        }
        try {
            findQuadraticEquationRoot(firstNumber, secondNumbner, thirtNumber);
        } catch (Exception e) {
            System.out.println("Found issue in calculation");
            e.getStackTrace();
        }

    }

    public static void findQuadraticEquationRoot(double a, double b, double c) {

        double sqrt_val = b * b - (4 * a * c);
        double d = sqrt(abs(sqrt_val));
        double baseValue = 2 * a;
        double valueOf_B = -b;

        if (d == 0) {
            System.out.println("Roots are real and same");
            System.out.println(valueOf_B / baseValue);
        }
        if (d < 0) {
            System.out.println("Roots are complex");
            System.out.println("first root is " + (valueOf_B + d) / baseValue + " second root is " + (valueOf_B - d) / baseValue);
        }
        if (d > 0) {
            System.out.println("Roots are real but not same ");
            System.out.println("first root is " + (valueOf_B + d) / baseValue + " second root is " + (valueOf_B - d) / baseValue);
        }

    }
}


