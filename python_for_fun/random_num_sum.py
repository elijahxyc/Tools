import random


MAX = 1000
output_num_1 = []
output_num_2 = []
left = 0

#算法1：直接计算，这样的问题是容易产生很多的0
for i in range(MAX):
    random_num = random.randint(0,MAX-left)
    output_num_1.append(random_num/1000)
    left = random_num + left

#算法2：分段计算：这样的好处是0可以变少，但是计算量会增加
jump_num = 100 #这个数字需要可以被1000整除
MAX_2 = int(MAX/jump_num)
print(MAX_2)

for j in range(jump_num):
    left = 0
    for i in range(MAX_2):
        random_num = 0
        if i == MAX_2 - 1:
            random_num = MAX_2 - left
        else:
            random_num = random.randint(0, MAX_2 - left)
        output_num_2.append(random_num)
        left = random_num + left

for i in range(len(output_num_2)):
    output_num_2[i] =  output_num_2[i]/1000

print(output_num_1)
print(output_num_2)