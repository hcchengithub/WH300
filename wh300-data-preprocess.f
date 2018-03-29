
    include excel.f 

    <o> 
    <h2>WKS RD 庫房月報表 over300d 將「手工標注」導回總表</h2>
    <h4><input type=checkbox></input> 用本程式的 excel （在 Task Bar 上閃爍的那個） 打開庫房的帳表。</h4>
    <h4><input type=checkbox></input> 對 Aging 欄排序，最老的在最上面。</h4>
    <h4><input type=checkbox></input> 若有新資料，請修改「excel帳表的參數.f」檔。</h4>
    <h4>&nbsp;&nbsp;&nbsp;&nbsp;帳表Excel參數設定檔 <input id=initFile type=text size=80% value="c:\Users\hcche\Documents\GitHub\WH300\excel帳表的參數.f"></h4>
    <h4>&nbsp;&nbsp;&nbsp;&nbsp;命令 "column . space row . cr" 可得 activeCell 的座標。
    <h4><input type=checkbox></input> 按下「執行」約需兩分半鐘的時間。</h4>
    <h4><input type=checkbox></input> 此時「總表」已獲 Tag 欄位。保存 excel 帳表。</h4>
	<h4><input type=button onclick="vm.execute('Labeled資料表>倒填回總表')" 
		value="執行"
		style="width:120px;height:40px;font-size:20px;margin-left:50px;"></h4>
    </o> er drop

    \ excel 帳表的參數 從外面手動設定
    
    char over300d    value 樞紐分析表名 // ( -- str ) worksheet name 英文大小寫不分
    char Detail      value 總表名       // ( -- str ) worksheet name 英文大小寫不分
    js> [1,1,23,271] value 樞紐分析表「整表」座標   // ( -- array ) [左上col,row,右下col,row]
    js> [5,4,22,270] value 樞紐分析表「資料」座標   // ( -- array ) [左上col,row,右下col,row]
    3                value 樞紐分析表「部門代碼row」 // ( -- int )
    3                value 樞紐分析表「料號column」 // ( -- int )
    char u           value 總表Tag欄     // ( -- int ) 欄位英文字母,大小寫不分
    char c           value 總表PartNo欄  // ( -- int ) 欄位英文字母,大小寫不分
    char n           value 總表部門欄    // ( -- int ) 欄位英文字母,大小寫不分

    \ 程式內部常數、變數
    
    0                value 總表部門欄#    // ( -- int )
    0                value 總表Tag欄#     // ( -- int )
    0                value 總表PartNo欄#  // ( -- int )
    {}               value 總表           // ( -- worksheet )
    {}               value 樞紐分析表     // ( -- worksheet )
    0                value Labeled資料表長度 // ( -- int )
    0                value partNo欄的長度 // ( -- int )
    {}               value Labeled資料表  // ( -- hash ) 料號、部門、數量
    ""               value csv            // ( -- csv ) Labeled資料表's csv string

    : 「樞紐分析表」手工Tag好的結果轉換成「Labeled資料表」 
        
        樞紐分析表名 activeWorkbook :> worksheets(pop()) to 樞紐分析表
        樞紐分析表 :: activate() \ 切換到「樞紐分析表」
        
        ( 參考座標在原點上 ) 1 1 goto <js>
        var et = vm.v('樞紐分析表「整表」座標')         // entire table    
        var da = vm.v('樞紐分析表「資料」座標')         // data area   
        var st_row = vm.v('樞紐分析表「部門代碼row」')  // The row of section ID
        var pn_col = vm.v('樞紐分析表「料號column」')   // The column of partNo
        var qty,partn,section, 料號=[], 部門=[], 數量=[]
        var csv = "料號,部門,數量,\n"
        for (var row=et[1]; row<=et[3]; row++){
            for (var col=et[0]; col<=et[2]; col++){
                var cell = push(col-1).push(row-1).execute('offset').pop()
                var inDA = col >= da[0] && col <= da[2] && row >= da[1] && row <= da[3]  
                if (cell.value) {
                    if (inDA && cell.Interior.Color!=16777215){  // Not white means Tagged
                        qty = cell.value 
                        section = push(cell.column-1).push(st_row-1).execute('offset').pop().value 
                        partn = push(pn_col-1).push(cell.row-1).execute('offset').pop().value 
                        csv += partn + ',' + section + ',' + qty + ',\n'  
                        料號.push(partn)
                        部門.push(section)
                        數量.push(qty)
                    }
                }
            }
        }
        push({'料號':料號,'部門':部門,'數量':數量}).dictate("to Labeled資料表")
        push(csv).dictate("to csv")
        </js>  
        
        \ 檢查一下
        Labeled資料表 :> 料號.length \ 55 OK 
        Labeled資料表 :> 部門.length \ 55 OK 
        Labeled資料表 :> 數量.length \ 55 OK
        ( a b c ) over ( a b c b ) = ( a b b=c ) -rot 
        ( b=c a b ) over ( b=c a b a ) = ( b=c a a=b )
        rot ( a a=b b=c ) and 
        if to Labeled資料表長度
        else ." Error! Labeled資料表 內部長度不一致 " cr then
        ;
    
    : 跳到「總表PartNo欄開頭」 ( -- ) \ 參考座標
        總表PartNo欄 1 + jump ;

    : Labeled資料表>倒填一項回總表 ( i -- ) \ i 是 Labeled資料表 的 index 
        跳到「總表PartNo欄開頭」 ( 參考座標 ) 
        <js>    
            index  = pop() // 外面 指定 Labeled資料表 某一 index 進來
            partNo = vm.v('Labeled資料表')['料號'][index]
            DEPT   = vm.v('Labeled資料表')['部門'][index]
            QTY    = vm.v('Labeled資料表')['數量'][index]
            
            for (row=1; row<=vm.v('partNo欄的長度'); row++){
                push(1).execute('nap') // 避免 js host 以為程式咬死而發警告干擾
                var cell = push(0).push(row-1).execute('offset').pop()
                var pn = cell.value
                var dp = push(vm.v('總表部門欄#')-vm.v('總表PartNo欄#'))
                         .push(row-1)
                         .execute('offset').pop().value
                         // 有可能是 undefined 表示還在庫房裡，未領用。
                if (pn==partNo && dp==DEPT){
                    tag_cell = 
                        push(vm.v('總表Tag欄#')-vm.v('總表PartNo欄#'))
                        .push(row-1)
                        .execute('offset').pop()
                        // 有可能是 undefined
                    tag_cell.value = 1
                    tag_cell.Interior.Color = 0x0000FF // Red 
                    QTY -= 1
                    if (QTY) continue
                    else break
                }
            }
        </js> ;

    : Labeled資料表>倒填回總表 ( -- ) \ 全部做完
        \ 參數 init 
            js> initFile.value readTextFile js: vm.dictate(pop()) 
            總表部門欄   :> toUpperCase() letter>column# to 總表部門欄#
            總表Tag欄    :> toUpperCase() letter>column# to 總表Tag欄#
            總表PartNo欄 :> toUpperCase() letter>column# to 總表PartNo欄#
        \ 取得手工標好的結果
            ." Start  " now . cr
            ." 「樞紐分析表」手工Tag好的結果轉換成「Labeled資料表」 " cr
            「樞紐分析表」手工Tag好的結果轉換成「Labeled資料表」 
            ." End  " now . cr 
        \ 「Labeled資料表」導回總表
            \ 切到總表 
            總表名 activeWorkbook :> worksheets(pop()) to 總表
            總表 :: activate() 跳到「總表PartNo欄開頭」
            column column#>letter dup char : + swap + ( c:c column )
            activeSheet :> range(pop()) bottom to partNo欄的長度 // ( -- int )

        \ 開始執行
            manual \ 讓 excel 不要自動 re-calculate 節省時間
            cr ." Start  " now . cr
            Labeled資料表長度 for 
                Labeled資料表長度 r@ - dup . space 
                ." Labeled資料表>倒填一項回總表 " cr
                Labeled資料表>倒填一項回總表 
            next 
            ." End  " now . cr 
            auto \ 恢復 excel 自動 re-calculate
        ;
        
