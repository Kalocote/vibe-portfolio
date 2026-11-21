age = input("你多大了？")

try:
	age = int(age)
	print("明年你就",age + 1)

except ValueError as e:
	print("出错了，具体是:",e)
