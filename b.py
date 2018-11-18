# def read_file(file_name):
#     with open(file_name) as f:
#         data = f.read()
#
#     print(data)
#
# def write_file(file_name):
#     with open(file_name, mode="w", encoding="utf-8") as f:
#         f.write("中国")
#
#
# f_path = r"E:\test\test.txt"
# # read_file(f_path)
# write_file(f_path)

# def test1(*args, **kwargs):
#     print(args)
#     print(kwargs)
#
# a=[1,2,3]
# b={"a":1, "b": 2, "c": 3}
#
# test1(*a, **b)
# test1((1, 2, 3), a=1, b=2, c=3)
# import traceback
#
# try:
#     print(1/0)
# except Exception as e:
#     print(traceback.format_exc())


# import os
# f_path = r"E:\test\new"
# if os.path.exists(f_path):
#     print("ok")
# else:
#     os.makedirs(f_path)
#     print("create oik")


# a="{name}@@{:.2f}@@{addr}".format(
#     name="china",
#     age=18,
#     addr="asdf"
# )

# b=a.split('@@')
# print(b)

a="{name}@@{age:.2f}@@{man:.2f}".format(
    name="ccc",
    age=12,
    man=15
)
import time
print(time.strftime("%Y/%m/%d"))

# a=[11,22,33]
# # a.pop(2)
# # print(a)
#
# for j, k in enumerate(a, 1):
#     print(j, k)
