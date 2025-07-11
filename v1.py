# 2024년도에 만든 교실 자리 뽑기 프로그램 V1
import random
import time

print("\n"+"교실 자리 뽑기 프로그램"+"\n")
print("==================="+"\n")
while True:
    print("교실 세로 줄을 몇 줄로 할 것인가요? (숫자만 작성)")
    print("선택 가능 : 4줄 / 5줄")
    line = input()
    print("==================="+"\n")
    if line == "4" or line == "5":
        print("[ "+line+"줄 ] 로 선택하셨습니다!"+"\n")
        while True:
            number = input("학급 전체 인원을 알려주세요."+"\n")
            if not number.isdigit():
                print("올바르지 않습니다, 다시 입력해주세요."+"\n")
            elif number.isdigit():
                number = int(number)
                if number > 50:
                    print("50명 이하로 다시 작성하세요."+"\n")
                else:
                    rewhile = True
                    while rewhile:
                        print(number,"명으로 실행하겠습니까? (Y / N)")
                        Start = input()
                        if Start == "Y" or Start == "y":
                            print("\n"+"============교탁============"+"\n")
                            number_list = list(range(1, number + 1))
                            random.shuffle(number_list)
                            for i in range(number):
                                print(f'{number_list[i]:<8.0f}', end=' ')
                                if (i + 1) % int(line) == 0:
                                    time.sleep(2)
                                    print()
                            print()
                            print("\n"+"다시 돌리기 (Y / N)")
                            reinput = input()
                            if reinput == "Y"or reinput == "y":
                                rewhile = True
                            else:
                                print("\n"+"프로그램을 종료합니다.")
                                break
                        elif Start == "N" or Start == "n":
                            print("\n"+"프로그램을 종료합니다.")
                            break
                        else:
                            print("올바르지 않습니다, 다시 입력해주세요."+"\n")
                    break
        break
    else:
        print("올바르지 않습니다, 다시 입력해주세요."+"\n")
