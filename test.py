def find_second_star_index(text):
    first_star_index = text.find('*')
    if first_star_index == -1:
        return -1  # 如果没有找到第一个星号，返回-1
    
    second_star_index = text.find('*', first_star_index + 1)
    return second_star_index

# 示例
input_string = "这是一个*示例字符串*，包含*多个*星号。"
second_star_index = find_second_star_index(input_string)
print(second_star_index)  # 输出: 8
print(input_string[second_star_index:])
