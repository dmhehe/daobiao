import os
import json



g_namespace = {}


def exec_and_return(expression: str):
    import simui.simwin
    global g_namespace
    
    g_namespace["win_w"] = simui.simwin.g_win_width
    # 将表达式分割为多个部分
    parts = expression.split(';')

    last_str = parts[-1]
    
    if last_str.find("=") == -1:
        parts[-1] = "__result = " + last_str
    else:
        new_line = "__result = " + last_str.split("=")[0]
        parts.append(new_line)
    
    new_expression = ";".join(parts)

    exec(new_expression, g_namespace)

    # 计算并返回最后一个表达式的结果
    return g_namespace["__result"]

# 使用函数
# result = execute_and_return("a = 5; b = 10; a + b")
# print(result)
