#include <stdio.h>
#include <windows.h>

int main() {
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);

    float scores[5];  // 存储5个成绩
    float sum = 0.0, average;
    int i;

    printf("请输入5个学生的成绩：\n");
    
    // 循环输入5个成绩
    for (i = 0; i < 5; i++) {
        printf("第%d个学生的成绩: ", i + 1);
        scanf("%f", &scores[i]);
        sum += scores[i];  // 累加总和
    }

    // 计算平均成绩
    average = sum / 5;

    // 输出结果
    printf("\n这5个学生的平均成绩为: %.2f\n", average);

    return 0;
}