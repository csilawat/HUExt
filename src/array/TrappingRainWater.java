package array;

public class TrappingRainWater {

    static int arr[] = new int[]{7, 4, 0, 9};

    public static void main(String[] args) {

        int totalWaterTrappedCount = trap(arr);
        System.out.println("Water can be trapped between the blocks during the rainy season is " + totalWaterTrappedCount);
    }

    public static int trap(int[] heights) {
        int leftSide = 0, rightSide = heights.length - 1, waterTrapped = 0;

        int maxLeft = heights[leftSide];
        int maxRight = heights[rightSide];

        while (leftSide < rightSide) {
            if (heights[leftSide] <= heights[rightSide]) {
                leftSide++;
                maxLeft = Integer.max(maxLeft, heights[leftSide]);
                waterTrapped += (maxLeft - heights[leftSide]);
            } else {
                rightSide--;
                maxRight = Integer.max(maxRight, heights[rightSide]);
                waterTrapped += (maxRight - heights[rightSide]);
            }
        }

        return waterTrapped;
    }
}
