
import simui
from simui import sim_class, sim_meth, sim_func, init_sim_global, show_popup

import main







#生成json的文件夹，项目里面没有用的，给看看参考而已
g_JsonPath = "D:/daobiao/json"

#生成ts声明代码的地方，项目有用的，因为是声明，上线编译成js就会没有掉的了
g_TsPath = "D:/daobiao/ts"

#bin文件就是所有json的zip压缩算法的压缩包，项目有用的， 要放到项目daobiao文件夹里面 
g_PackPath = "D:/daobiao/pack"

#xlsx的文件夹
g_XlsxFloderPath = "D:/daobiao/xlsx"


attr_list2 = [
    ["g_JsonPath", "json的文件夹", "str", {"default": ""}],
    ["g_TsPath", "ts的文件夹", "str", {"default": ""}],
    ["g_PackPath", "pack的文件夹", "str", {"default": ""}],
    ["g_XlsxFloderPath", "xlsx的文件夹", "str", {"default": ""}],
]


gg = init_sim_global("gg", attr_list2)




@sim_func("干活！", {"iPrior": 0})
def do():
    print("干活！！！！！！！！！！！！")
    
    if not gg.g_JsonPath:
        show_popup("请设置json的文件夹")
        return
    
    if not gg.g_TsPath:
        show_popup("请设置ts的文件夹")
        return
    
    
    if not gg.g_PackPath:
        show_popup("请设置pack的文件夹")
        return
    
    if not gg.g_XlsxFloderPath:
        show_popup("请设置xlsx的文件夹")
        return
    
    main.g_JsonPath = gg.g_JsonPath
    main.g_TsPath = gg.g_TsPath
    main.g_PackPath = gg.g_PackPath
    main.g_XlsxFloderPath = gg.g_XlsxFloderPath
    
    main.main()
    show_popup("完成！！！")

simui.show_ui("daobiao")
