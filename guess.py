import random
answer = random.randint(1, 100)
guess = 0
tries = 0

while guess != answer:
    try:
        guess = int(input("猜一个1-100的数字："))
    except ValueError:
        print("请输入数字！")
        continue

    if guess < answer:
        print("小了")
    elif guess > answer:
        print("大了")

    tries = tries + 1

print("猜对了！你用了",tries,"次")
