package loops;

public class PrintPattern {

    public static void main(String[] args) {

        int Size = 4;
        printPatternWithGivenSize(Size);

    }

    static void printPatternWithGivenSize(int N) {

        int index, spaceIndex;
        for (index = 0; index < N; index++) {
            System.out.print("*");

            for (spaceIndex = 0; spaceIndex < 2 * (index); spaceIndex++)
                System.out.print(" ");

            if (index != 0 && index != N - 1)
                System.out.print("*");
            else
                System.out.print(" ");

            for (spaceIndex = 0; spaceIndex < 2 * (N - index - 1); spaceIndex++)
                System.out.print(" ");

            System.out.print("*");

            System.out.println(" ");
        }
    }
}
