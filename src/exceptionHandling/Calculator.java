package exceptionHandling;

import java.util.InputMismatchException;
import java.util.Scanner;

public class Calculator {

    static long firstNumber, secondNumber;
    static boolean condition = true;

    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);

        String operation;
        String menu = "Choose an operation:\n+ Addition\n- Subtract\n* Multiply\n/ Divide\nroot Square Root\n Exit";
        while (condition == true) {
            menu(menu);
            operation = scanner.next();

            try {
                switch (operation) {
                    case "+":
                        System.out.println("Enter first number : ");
                        firstNumber = scanner.nextInt();
                        System.out.println("Enter second number : ");
                        secondNumber = scanner.nextInt();
                        add(firstNumber, secondNumber);
                        break;
                    case "-":
                        System.out.println("Enter first number : ");
                        firstNumber = scanner.nextInt();
                        System.out.println("Enter second number : ");
                        secondNumber = scanner.nextInt();
                        subtract(firstNumber, secondNumber);
                        break;
                    case "*":
                        System.out.println("Enter first number : ");
                        firstNumber = scanner.nextInt();
                        System.out.println("Enter second number : ");
                        secondNumber = scanner.nextInt();
                        multiply(firstNumber, secondNumber);
                        break;
                    case "/":
                        System.out.println("Enter first number : ");
                        firstNumber = scanner.nextInt();
                        System.out.println("Enter second number : ");
                        secondNumber = scanner.nextInt();
                        divide(firstNumber, secondNumber);
                        break;
                    case "root":
                        System.out.println("Enter number : ");
                        firstNumber = scanner.nextInt();
                        sqrt(firstNumber);
                        break;
                    case "Exit":
                        System.out.println("You have exited from the calculator");
                        System.exit(0);
                        condition = false;
                        break;
                    default:
                        System.out.println("you have selected Invalid operation");
                        System.out.println("please select right operation");
                }
            } catch (InputMismatchException e) {

                System.out.println("Your enter number is not a valid number");
                System.out.println("Please enter a valid number");
                System.out.println("Or type Exit");

            }
        }

        scanner.close();
    }

    public static void menu(String menu) {
        System.out.println(menu);
    }

    public static void add(long firstNumber, long secondNumber) {

        System.out.println(firstNumber + secondNumber);
    }

    public static void subtract(long firstNumber, long secondNumber) {
        System.out.println(firstNumber - secondNumber);
    }

    public static void multiply(long firstNumber, long secondNumber) {
        System.out.println(firstNumber * secondNumber);
    }

    public static void divide(long firstNumber, long secondNumber) {
        if (secondNumber != 0) {
            System.out.println(firstNumber / secondNumber);
        } else {
            System.out.println("Don't enter Second no as 0 (Zero), Cannot divide by 0");
        }
    }

    public static void sqrt(long firstNumber) {
        System.out.println(Math.sqrt(firstNumber));
    }
}
