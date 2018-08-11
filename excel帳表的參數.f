
    \ excel 帳表的參數

    char Over300D    to 樞紐分析表名 \ ( -- str ) worksheet name 英文大小寫不分
    char Detail      to 總表名       \ ( -- str ) worksheet name 英文大小寫不分
    js> [4,7,22,258] to 樞紐分析表「資料」座標    \ ( -- array ) [左上col,row,右下col,row] 用 column . space row . cr 命令可得 activeCell 的座標
    5                to 樞紐分析表「部門代碼row」 \ ( -- int )
    char b           to 樞紐分析表「料號column」  \ ( -- int )
    char t           to 總表Tag欄    \ ( -- int ) 欄位英文字母,大小寫不分
    char c           to 總表PartNo欄 \ ( -- int ) 欄位英文字母,大小寫不分
    char m           to 總表部門欄   \ ( -- int ) 欄位英文字母,大小寫不分

    \ ------- the END ------