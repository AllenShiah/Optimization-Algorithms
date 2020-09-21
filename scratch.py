import pandas as pd
import numpy as np

# 從 excel 中讀取資料
def importdata(path):
    global data  # 把匯入資料儲存在全域變數
    sheet_name="job"    # 表單名稱
    data = pd.read_excel(path,sheet_name,index_col=0)   # pandas 讀取資料

# 調換 Job order
def swap(job_order):
    global neighbors,neighbors_set,data,possible_tabu
    for i in range(len(job_order)-1):
        job_order[i],job_order[i+1]=job_order[i+1],job_order[i]   #調換 Job order 位置
        data=data[job_order]   # 更換 DataFrame 中的 Job Order 順序
        neighbors=np.vstack((job_order,neighbors))   # 每一輪的調換結果 array 儲存
        Objective_value(job_order)
        switch=job_order[i]+','+job_order[i+1]   # 調換過的 Job (需要再被調回才可以進入 Tabu list)
        switch_split=switch.split(',')
        neighbors_set=np.vstack((switch_split,neighbors_set))   # 調換過的 Job 在調整進入 Tabu list 前先儲存
        possible_tabu_change_order(switch_split)   #進入調換函數
        job_order[i], job_order[i + 1] = job_order[i + 1], job_order[i]   #調換回原來的 Job order 供下次次迴圈使用

# 轉換成正確可以進入 Tabu list 的值
def possible_tabu_change_order(switch_split):
    global possible_tabu
    switch_split[0],switch_split[1]=switch_split[1],switch_split[0]
    possible_tabu = np.vstack((switch_split, possible_tabu))

# 計算 Objective Value 的 function
def Objective_value(job_order):
    global round_score
    score=0
    Pj=0
    for y in job_order:
        Pj+=data[y][0]   # 每個 Job 的 Pj
        Dj=data[y][1]    # 每個 Job 的 Dj
        Wj=data[y][2]    # 每個 Job 的 Wj
        z=max(Pj-Dj,0)   # Tardiness of Job
        score+=z*Wj   #計算 Total weighted tardiness
    round_score=np.vstack((score,round_score))   # 把每一個 Job order 算出來的 Objective Value 存進 round_score array 裡

# 檢查 min(Job order) 中是否被 tabu
def compare_tabu_list(index,counter2):
    global tabulist,possible_tabu,neighbors_set,counter

    if counter2 == 0:
        counter+=1
        #print("ssssssssssssssssssss")
        return False
    else:
        for i in range(len(tabulist)):
            if len(tabulist)==2:
                tabulist=np.delete(tabulist,1,0)
            result=neighbors_set[index]==tabulist[i]
            if result[0] or result[1]==True:
                return True
            else:
                return False
            print(result)

# 主程式
if __name__=="__main__":
    # excel 匯入資料的儲存
    data = 0
    io = str(input("請輸入 excel 檔案讀取路徑: ")) + r"\tabu_search.xlsx"  # 輸入資料的 directory
    importdata(io)  # 啟動 importdata function
    print(data)  # 先印出總工作表的 Pandas DataFrame
    print("===========================")

    # 全域變數區
    data_array=np.array(data) #將匯入的 DataFrame 轉成 array
    job_order=data.columns.tolist()  # 演算法的起始值 (Series)
    neighbors=np.empty((0,4))   # 儲存每局可能的 neighbors
    neighbors_set=np.empty((0,2))   # 儲存 neighbors 調換的紀錄 (用來與 Tabu list 做比較)
    possible_tabu=np.empty((0,2))   # 調換正確順序之 neighbors_set (可以進入 Tabu list)
    round_score=np.empty((0,1))   # 儲存每一個 job order 的值
    tabulist=np.empty((0,2)) # tabulist 的儲存
    final_job_order=np.empty((0,4))
    iter_times=100   # 演算法遞迴次數
    aspiration_criterion = 0  # objective value 的值儲存
    indice=[]   # 將 min max middle index 的值組成一個 list 供迴圈讀取
    counter = 0

    # 主程式執行區
    for x in range(iter_times):
        if len(round_score)==3:   # 刪除前一個 interation 的紀錄，供下一個 interation 使用
            neighbors=np.empty((0,4))
            neighbors_set = np.empty((0, 2))
            round_score=np.empty((0,1))
            possible_tabu = np.empty((0, 2))
            print("---------------------------")

        swap(job_order)   # Tabu Search Algo starts from here
        round_score_list=round_score.tolist()   # 把此 round 的結果轉成 list 才可以用 index
        min_round_score_index = round_score_list.index(min(round_score_list))   # 取出每個 iteration 中最小的值 的 index
        max_round_score_index=round_score_list.index(max(round_score_list))   # 取出每個 iteration 中最大的值 的 index
        for y in range(len(round_score_list)):   # 取出每個 interation 中 中間值的 index
            if min(round_score_list)<round_score_list[y]<max(round_score_list):
                middle_round_score_index=y
        indice=[min_round_score_index,middle_round_score_index,max_round_score_index]   # 把 min max middle 的 index 放進去

        # 以下為將 min middle max 的 job order 依序與 tabulist 做比較，找出最適合的 Job order 為止
        for z in indice:

            compare_result=compare_tabu_list(z,counter)   # 依序將 min middle max 傳入 compare_tabu_list 看看是否有被禁忌

            if compare_result == False:
                job_order=neighbors[z]
                aspiration_criterion=round_score[z]
                tabulist = np.vstack((possible_tabu[z], tabulist))
                break
        print("Tabu list: ")
        print(tabulist)
        print("此局的最佳排序: ",job_order)
        print("aspiration criterion: ",aspiration_criterion)
    print("===========================")
    print("Conclusion")
    print("當前工作最佳排列:",job_order)
    print("當前最佳解: ",aspiration_criterion)
    print("總共遞迴: ",iter_times," 次")