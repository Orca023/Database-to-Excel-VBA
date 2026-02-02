## Microsoft Office Excel client to Database
#### 使用電子表格作爲用戶端訪問資料庫實現新增、刪除、修改、查找資料等操作.
#### Microsoft Office Excel client to Database ( Microsoft Access , MongoDB , MariaDB , etc ) operate CRUD ( Create , Read , Update , Delete ).
#### Microsoft Office Excel Professional 2019 x86_64
#### Microsoft Access Professional 2019 x86_64
#### MongoDB , MariaDB

---

<p word-wrap: break-word; word-break: break-all; overflow-x: hidden; overflow-x: hidden;></p>

項目使用了微軟電子表格 ( Windows - Office - Excel - Visual Basic for Applications ) 應用的第三方擴展 :  `clsBrowser.cls` , `clsCore.cls` , `clsJsConverter.cls` 三個類模組，由 codeproject 網站發佈取得.

[第三方擴展類模組提供網站 codeproject 裏的 Automate Chrome or Edge using VBA 庫 ( Tips ) 官方説明頁](https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA): 
https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA

項目使用了微軟電子表格 ( Windows - Office - Excel - Visual Basic for Applications ) 應用的第三方擴展 :  `JsonConverter.bas` 模組，由 GitHub 網站倉庫 ( Repository ) : VBA-JSON 發佈取得.

[相互轉換 JSON 字符串與 Excel-VBA-Dict 對象 ( Object ) 使用的第三方擴展類模組 VBA-JSON 官方 GitHub 網站倉庫頁](https://github.com/VBA-tools/VBA-JSON): 
https://github.com/VBA-tools/VBA-JSON.git

---

一. 確保 Microsoft Window11 系統的 Access 資料庫已安裝配置成功，和 MongoDB , MariaDB 資料庫的伺服端應用 ( Server ) 已安裝配置成功且已啓動運行，啓動 Microsoft Office Excel 應用.

二. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/CDPimport/` 裏的 Microsoft Excel VBA 類模組 ( Class Modul ) : `clsBrowser.cls` , `clsCore.cls` , `clsJsConverter.cls`

三. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/` 裏的 Microsoft Excel VBA 窗體 ( Form ) : `DatabaseControlPanel.frm` , `DatabaseControlPanel.frx`

四. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/` 裏的 Microsoft Excel VBA 模組 ( Module ) : `DatabaseDispatchModule.bas`

五. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/` 裏的 Microsoft Excel VBA 模組 ( Module ) : `DatabaseModule.bas`

六. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/` 裏的 Microsoft Excel VBA 模組 ( Module ) : `DatabaseMongoDB.bas`

七. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/` 裏的 Microsoft Excel VBA 模組 ( Module ) : `DatabaseMariaDB.bas`

八. 手動操作 Microsoft Excel 應用, 載入文件夾 `./Database-to-Excel-VBA/` 裏的 Microsoft Excel VBA 對象 ( Object ) : `ThisWorkbook.cls`

九. 啓動 MongoDB 資料庫的伺服端應用 ( Server ) 伺服器 :
```
`C:\Database-to-Excel-VBA\MongoDB> C:/Database-to-Excel-VBA/MongoDB/Server/8.2/bin/mongod.exe --config=C:/Database-to-Excel-VBA/MongoDB/NodejsToMongoDB/mongod.cfg`
```

十. 啓動自定義的操作 MongoDB 資料庫的 Node.js 伺服器 :
```
`C:\Database-to-Excel-VBA\MongoDB> C:/Database-to-Excel-VBA/Nodejs/Nodejs-22.20.0/node.exe C:/Database-to-Excel-VBA/MongoDB/NodejsToMongoDB/Nodejs2MongodbServer.js host=::0 port=27016 number_cluster_Workers=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1`
```

十一. 啓動 MariaDB 資料庫的伺服端應用 ( Server ) 伺服器 :
```
`C:\Database-to-Excel-VBA\MariaDB> C:/Database-to-Excel-VBA/MariaDB/MariaDB10.11/bin/mysqld.exe`
```

十二. 啓動自定義的操作 MariaDB 資料庫的 Python 伺服器 :
```
C:\Database-to-Excel-VBA\MariaDB> C:/Database-to-Excel-VBA/MariaDB/PythonToMariaDB/Scripts/python.exe C:/Database-to-Excel-VBA/MariaDB/PythonToMariaDB/src/Python2MariaDBServer.py host=::0 port=27016 Is_multi_thread=False number_Worker_process=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
```
或者 : 
```
C:\Database-to-Excel-VBA\MariaDB> C:/Database-to-Excel-VBA/Python/Python311/python.exe C:/Database-to-Excel-VBA/MariaDB/PythonToMariaDB/src/Python2MariaDBServer.py host=::0 port=27016 Is_multi_thread=False number_Worker_process=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
```

十三. 運行 Microsoft Office Excel VBA 宏擴展應用 : `Database-to-Excel-VBA` 選擇 `operation panel` 選項, 從 Microsoft Excel 應用的「`加載項 ( add-in )`」菜單裏, 選擇 : 「 `Manipulate database` 」 → 「 `operation panel` 」, 加載顯示 `operation panel` 人機交互介面.

十四. 測試 Microsoft Office Excel VBA 宏擴展應用 : `Database-to-Excel-VBA` 項目使用電子表格 ( Microsoft Office Excel ) 鏈接 `Microsoft Access , MongoDB , MariaDB` 等資料庫 , 通過讀取在電子表格 ( Microsoft Office Excel ) 指定位置的操作指令，實現新增 ( Create )、刪除 ( Delete )、修改 ( Update )、查找 ( Read ) 資料等操作, 將讀取結果存儲在電子表格 ( Microsoft Office Excel ) 指定位置.

---

項目空間裏的電子表格 Microsoft Office Excel 檔 : 「 `Database.xlsm` 」 已經載入 :

第三方類模組 ( Class Modul ) : `clsBrowser.cls` , `clsCore.cls` , `clsJsConverter.cls`

窗體 ( Form ) : `DatabaseControlPanel.frm` , `DatabaseControlPanel.frx`

模組 ( Module ) : `DatabaseDispatchModule.bas` , `DatabaseModule.bas` , `DatabaseMongoDB.bas` , `DatabaseMariaDB.bas`

對象 ( Object ) : `ThisWorkbook.cls`

可直接從 Microsoft Office Excel 應用的「`加載項 ( add-in )`」菜單裏, 選擇 : 「 `Manipulate database` 」 → 「 `operation panel` 」, 加載顯示 `operation panel` 人機交互介面.

之後即可測試使用電子表格 ( Microsoft Office Excel VBA ) 鏈接 `Microsoft Access , MongoDB , MariaDB` 等資料庫 , 通過讀取在電子表格 ( Microsoft Office Excel ) 指定位置的操作指令，實現新增 ( Create )、刪除 ( Delete )、修改 ( Update )、查找 ( Read ) 資料等操作, 將讀取結果存儲在電子表格 ( Microsoft Office Excel ) 指定位置.

---

使用微軟電子表格 Microsoft Office Excel VBA 鏈接 `Microsoft Access , MongoDB , MariaDB` 等資料庫 , 通過讀取在電子表格 ( Microsoft Office Excel ) 指定位置的操作指令，實現新增 ( Create )、刪除 ( Delete )、修改 ( Update )、查找 ( Read ) 資料等操作説明 :

1. 項目架構執行序 :

   1). 啓動 Microsoft Office Excel Professional 2019 應用, 電子表格 Excel 應用會自動運行已載入的模組 ( Module ) 和類模組 ( Class Modul ), 其中載入的調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 裏的自定義過程 ( Subroutine ) : `MenuSetup()` 會修改電子表格 Excel 的菜單欄 ( Menu Bar ) 向「 `加載項 ( add-in )` 」菜單下寫入自定義的 Microsoft Excel VBA 宏擴展應用 : `Database-to-Excel-VBA` 標簽.

   2). 單擊電子表格 Excel 「 `加載項 ( add-in )` 」菜單 ( Menu ) 下 Microsoft Excel VBA 宏擴展應用 : `Database-to-Excel-VBA` 標簽, 首先執行的是調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ).

   3). 調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 導入加載調用操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ), 並讀取操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) 裏的自定義配置參數值. 其中, 鏈接操控資料庫 Microsoft Access 的子過程 ( Sub ) 脚本代碼, 也存放在該操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏.

   同時，調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 導入加載調用操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ), 用於鏈接操控 MongoDB 資料庫.

   同時，調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 導入加載調用操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ), 用於鏈接操控 MariaDB 資料庫.

   4). 同時, 調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 調用窗體 ( Form ) 對象 ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ), 並根據操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏的自定義配置參數值, 爲窗體 ( Form ) 介面 ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 賦初值, 窗體 ( Form ) 對象 ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 是人機交互介面.

   5). 手動操控窗體 ( Form ) 介面, 自定義點選參數，啓動運行.

   6). 操控介面窗體 ( Form ) ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 調用, 操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏的 Sub Run 子過程, 啓動鏈接操控資料庫 ( Microsoft Access , MongoDB , MariaDB , etc ) 實現新增、刪除、修改、查找資料等操作.

   7). 操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏的 Sub Run 子過程, 讀取操控介面窗體 ( Form ) ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 裏自定義點選輸入的參數值.

   8). 操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏的 Sub Run 子過程, 根據操控介面窗體 ( Form ) ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 裏自定義點選輸入的參數值, 選擇執行如下後續動作 :

      8.1). 調用操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏的 Sub Run_Access 子過程, 其中包括如下動作 :

         8.1.1). 從自定義傳入的電子表格 Excel 指定位置, 讀取待上傳或操控資料庫 Microsoft Access 的資訊.

         8.1.2). 引用第三方類模組 ( Class Modul ) : ( `./Database-to-Excel-VBA/CDPimport/clsBrowser.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsCore.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsJsConverter.cls` ) , 轉換待上傳的資訊格式 ( 例如, 將二維數組 ( Array 2 Dimension ) 類型的數據轉換爲 JSON 字符串類型的數據 ) , 使之可以被資料庫 Microsoft Access 軟體識別處理, 從而達到寫入資料庫 Microsoft Access 目的.

         8.1.3). 執行鏈接操控資料庫 ( Microsoft Access ) 實現新增、刪除、修改、查找資料等操作.

         8.1.4). 讀取從資料庫 Microsoft Access 返回的資訊.

         8.1.5). 引用第三方類模組 ( Class Modul ) : ( `./Database-to-Excel-VBA/CDPimport/clsBrowser.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsCore.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsJsConverter.cls` ) , 轉換從資料庫 Microsoft Access 返回的資訊格式 ( 例如, 將 JSON 字符串類型的數據轉換爲二維數組 ( Array 2 Dimension ) 類型的數據 ) , 使之可以被電子表格 Excel 軟體識別, 從而實現寫入電子表格 Excel 目的.

         8.1.6). 將從資料庫 Microsoft Access 返回的資訊, 寫入自定義傳入的電子表格 Excel 指定位置.

      8.2). 調用操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) 裏的 Sub Run_MongoDB 子過程, 其中包括如下動作 :

         8.2.1). 從自定義傳入的電子表格 Excel 指定位置, 讀取待上傳或操控資料庫 MongoDB 的資訊.

         8.2.2). 引用第三方類模組 ( Class Modul ) : ( `./Database-to-Excel-VBA/CDPimport/clsBrowser.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsCore.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsJsConverter.cls` ) , 轉換待上傳的資訊格式 ( 例如, 將二維數組 ( Array 2 Dimension ) 類型的數據轉換爲 JSON 字符串類型的數據 ) , 使之可以被資料庫 MongoDB 軟體識別處理, 從而達到寫入資料庫 MongoDB 目的.

         8.2.3). 執行鏈接操控資料庫 ( MongoDB ) 實現新增、刪除、修改、查找資料等操作.

         8.2.4). 讀取從資料庫 MongoDB 返回的資訊.

         8.2.5). 引用第三方類模組 ( Class Modul ) : ( `./Database-to-Excel-VBA/CDPimport/clsBrowser.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsCore.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsJsConverter.cls` ) , 轉換從資料庫 MongoDB 返回的資訊格式 ( 例如, 將 JSON 字符串類型的數據轉換爲二維數組 ( Array 2 Dimension ) 類型的數據 ) , 使之可以被電子表格 Excel 軟體識別, 從而實現寫入電子表格 Excel 目的.

         8.2.6). 將從資料庫 MongoDB 返回的資訊, 寫入自定義傳入的電子表格 Excel 指定位置.

      8.3). 調用操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ) 裏的 Sub Run_MariaDB 子過程, 其中包括如下動作 :

         8.3.1). 從自定義傳入的電子表格 Excel 指定位置, 讀取待上傳或操控資料庫 MariaDB 的資訊.

         8.3.2). 引用第三方類模組 ( Class Modul ) : ( `./Database-to-Excel-VBA/CDPimport/clsBrowser.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsCore.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsJsConverter.cls` ) , 轉換待上傳的資訊格式 ( 例如, 將二維數組 ( Array 2 Dimension ) 類型的數據轉換爲 JSON 字符串類型的數據 ) , 使之可以被資料庫 MariaDB 軟體識別處理, 從而達到寫入資料庫 MariaDB 目的.

         8.3.3). 執行鏈接操控資料庫 ( MariaDB ) 實現新增、刪除、修改、查找資料等操作.

         8.3.4). 讀取從資料庫 MariaDB 返回的資訊.

         8.3.5). 引用第三方類模組 ( Class Modul ) : ( `./Database-to-Excel-VBA/CDPimport/clsBrowser.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsCore.cls` ) , ( `./Database-to-Excel-VBA/CDPimport/clsJsConverter.cls` ) , 轉換從資料庫 MariaDB 返回的資訊格式 ( 例如, 將 JSON 字符串類型的數據轉換爲二維數組 ( Array 2 Dimension ) 類型的數據 ) , 使之可以被電子表格 Excel 軟體識別, 從而實現寫入電子表格 Excel 目的.

         8.3.6). 將從資料庫 MariaDB 返回的資訊, 寫入自定義傳入的電子表格 Excel 指定位置.

2. 項目將自定義的操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ) 分別作爲獨立的一個模組 ( Module ) 設計, 目的是, 與調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 分開, 解耦合, 這樣便於日後維護擴展功能, 增加更多元的操控介面, 使之可選擇的, 適用於讀取更多目標網站頁面裏顯示的資訊.

   同樣的, 項目將調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 作爲獨立的一個模組 ( Module ) 設計, 與窗體 ( Form ) 對象 ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 分開, 其目的也是爲了, 解耦合, 便於日後維護擴展功能, 增加更多元的操控介面, 使之可選擇的, 適用於讀取更多目標網站頁面裏顯示的資訊.

   若不考慮日後的功能擴展, 可取消獨立的調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 設計, 將之全部功能, 整合入窗體 ( Form ) 對象 ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 裏, 這樣可降低項目架構的複雜性, 更易於理解.

3. 若想擴展功能, 增加更多元的操控介面, 使之可選擇的, 適用於讀取更多種資料庫 ( Database ) 軟體, 可新增複製模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) 或 ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ) 並自定義重新命名 , 修改模組 ( Module ) 裏的 Sub Run_MongoDB 或 Sub Run_MariaDB 子過程 , 根據需要自定義修改設計編寫代碼脚本即可, 這一操作的目的, 是爲實現新增一組操作介面的效果, 例如像 ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) 或 ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ) 類似的.

   并且, 需要修改調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 裏的代碼, 使其可以正確找到調用自定義擴展新增的操作模組 ( Module ) 並正確的讀取適配合規的自定義擴展新增的配置參數初值, 例如像 ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 類似的.

   并且, 需要修改操作模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseModule.bas` ) 裏的代碼, 使其可以正確找到調用自定義擴展新增的操作模組 ( Module ) 並正確的調用運行因應子過程 ( Sub ) 的脚本代碼, 例如像 ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) 或 ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ) 類似的.

   并且, 需要修改窗體 ( Form ) 對象 ( `./Database-to-Excel-VBA/DatabaseControlPanel.frx` ) , ( `./Database-to-Excel-VBA/DatabaseControlPanel.frm` ) 裏的代碼, 使其可以正確適配合規的顯示自定義擴展新增的模組 ( Module )的配置參數值, 作爲人機交互介面, 可以正確的操控自定義擴展新增的模組 ( Module ) 引用第三方類模組 ( Class Modul ) : `clsBrowser.cls` , `clsCore.cls` , `clsJsConverter.cls` 處理資料, 例如像 ( `./Database-to-Excel-VBA/DatabaseMongoDB.bas` ) 或 ( `./Database-to-Excel-VBA/DatabaseMariaDB.bas` ) 類似的.

4. 項目空間裏的文件夾 `MongoDB` 是一組使用計算機程式設計語言 ( JavaScript ) 解釋器 ( Node.js ) 自定義創建的 http 伺服器 ( Server ) 應用, 電子表格 Microsoft Excel VBA 直接訪問此程式設計語言 ( JavaScript ) 解釋器 ( Interpreter : Node.js ) 的 http 伺服器 ( Server ) 應用並向其發送指令, 然後, 此程式設計語言 ( JavaScript ) 解釋器 ( Interpreter : Node.js ) 的 http 伺服器 ( Server ) 應用, 再鏈接驅動資料庫 MongoDB 伺服器端軟體, 實現操控資料庫 ( MongoDB ) 新增、刪除、修改、查找資料等操作, 這樣設計目的是, 起到隔離電子表格 Microsoft Excel VBA 直連訪問資料庫 MongoDB 伺服器 ( Server ) 的作用.

   使用計算機程式設計語言 ( JavaScript ) 解釋器 ( Interpreter : Node.js ) 自定義創建的 http 伺服器 ( Server ) 應用, 運行需要 Node.js 解釋器 ( Interpreter ) 環境, 所以運行之前, 需對作業系統 ( Operating System ) 安裝配置 Node.js 解釋器 ( Interpreter ) 環境成功方可.

   可在 Linux-Ubuntu 系統的控制臺命令列人機交互介面窗口 ( Ubuntu-bash ) 使用如下指令, 安裝配置 Node.js 解釋器 ( Interpreter ) 環境 :
   ```
   root@localhost:~# sudo apt install nodejs
   root@localhost:~# sudo apt install npm
   ```
   可在 Linux-Ubuntu 系統的控制臺命令列人機交互介面窗口 ( Ubuntu-bash ) 使用如下指令, 啓動運行此計算機程式設計語言 ( JavaScript ) 解釋器 ( Interpreter : Node.js ) 創建的 http 伺服器 ( Server ) 應用 :
   ```
   root@localhost:~# /bin/node ./Database-to-Excel-VBA/MongoDB/NodejsToMongoDB/Nodejs2MongodbServer.js host=::0 port=27016 number_cluster_Workers=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
   ```

   另,

   項目調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 裏自定義的子過程 : `Sub runMongoDBServerSideApplication()` 是微軟電子表格 Microsoft Excel VBA 調用視窗 ( Windows ) 系統裏的 `WScript.Shell` 對象 ( Object ) 裏的 `.Exec` 方法創建子進程 ( child Process ) 並再調用視窗 ( Windows ) 系統裏的 shell 語句控制臺命令行 ( cmd.exe ) 執行 Bash 語句運行資料庫 MongoDB 伺服器 ( Server ) 端應用的二進位可執行檔 ( .exe ) 從而實現, 單擊 ( Click ) 微軟電子表格 Microsoft Excel 應用軟體的菜單欄 ( Menu bar ) 裏自定義的子菜單，即可一鍵快捷啓動資料庫 MongoDB 伺服器 ( Server ) 端應用的效果.

   項目調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 裏自定義的子過程 : `Sub runMongoDBhttpServer()` 是微軟電子表格 Microsoft Excel VBA 調用視窗 ( Windows ) 系統裏的 `WScript.Shell` 對象 ( Object ) 裏的 `.Exec` 方法創建子進程 ( child Process ) 並再調用視窗 ( Windows ) 系統裏的 shell 語句控制臺命令行 ( cmd.exe ) 執行 Bash 語句運行程式設計語言 ( JavaScript ) 解釋器 ( Interpreter : Node.js ) 的二進位可執行檔 ( .exe ) 加載執行自定義創建的 http 伺服器 ( Server ) 應用代碼脚本檔 ( .js ) 從而實現, 單擊 ( Click ) 微軟電子表格 Microsoft Excel 應用軟體的菜單欄 ( Menu bar ) 裏自定義的子菜單, 即可一鍵快捷啓動使用計算機程式設計語言 ( JavaScript ) 解釋器 ( Interpreter : Node.js ) 自定義創建的 http 伺服器 ( Server ) 應用的效果.

5. 項目空間裏的文件夾 `MariaDB` 是一組使用計算機程式設計語言 ( Python ) 自定義創建的 http 伺服器 ( Server ) 應用, 電子表格 Microsoft Excel VBA 直接訪問此程式設計語言 ( Python ) 的 http 伺服器 ( Server ) 應用並向其發送指令, 然後, 此程式設計語言 ( Python ) 的 http 伺服器 ( Server ) 應用, 再鏈接驅動資料庫 MariaDB 伺服器端軟體, 實現操控資料庫 ( MariaDB ) 新增、刪除、修改、查找資料等操作, 這樣設計目的是, 起到隔離電子表格 Microsoft Excel VBA 直連訪問資料庫 MariaDB 伺服器 ( Server ) 的作用.

   使用計算機程式設計語言 ( Python ) 自定義創建的 http 伺服器 ( Server ) 應用, 運行需要 Python 解釋器 ( Interpreter ) 環境, 所以運行之前, 需對作業系統 ( Operating System ) 安裝配置 Python 解釋器 ( Interpreter ) 環境成功方可.

   可在 Linux-Ubuntu 系統的控制臺命令列人機交互介面窗口 ( Ubuntu-bash ) 使用如下指令, 安裝配置 Python 解釋器 ( Interpreter ) 環境 :
   ```
   root@localhost:~# sudo apt install python3
   root@localhost:~# sudo apt install pip
   ```
   可在 Linux-Ubuntu 系統的控制臺命令列人機交互介面窗口 ( Ubuntu-bash ) 使用如下指令, 啓動運行計算機程式設計語言 ( Python ) 創建的 http 伺服器 ( Server ) 應用 :
   ```
   root@localhost:~# ./Database-to-Excel-VBA/MariaDB/PythonToMariaDB/Scripts/python ./Database-to-Excel-VBA/MariaDB/PythonToMariaDB/src/Python2MariaDBServer.py host=::0 port=27016 Is_multi_thread=False number_Worker_process=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
   ```
   或者 :
   ```
   root@localhost:~# /bin/python3 ./Database-to-Excel-VBA/MariaDB/PythonToMariaDB/src/Python2MariaDBServer.py host=::0 port=27016 Is_multi_thread=False number_Worker_process=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
   ```

   另,

   項目調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 裏自定義的子過程 : `Sub runMariaDBServerSideApplication()` 是微軟電子表格 Microsoft Excel VBA 調用視窗 ( Windows ) 系統裏的 `WScript.Shell` 對象 ( Object ) 裏的 `.Exec` 方法創建子進程 ( child Process ) 並再調用視窗 ( Windows ) 系統裏的 shell 語句控制臺命令行 ( cmd.exe ) 執行 Bash 語句運行資料庫 MariaDB 伺服器 ( Server ) 端應用的二進位可執行檔 ( .exe ) 從而實現, 單擊 ( Click ) 微軟電子表格 Microsoft Excel 應用軟體的菜單欄 ( Menu bar ) 裏自定義的子菜單，即可一鍵快捷啓動資料庫 MariaDB 伺服器 ( Server ) 端應用的效果.

   項目調度模組 ( Module ) ( `./Database-to-Excel-VBA/DatabaseDispatchModule.bas` ) 裏自定義的子過程 : `Sub runMariaDBhttpServer()` 是微軟電子表格 Microsoft Excel VBA 調用視窗 ( Windows ) 系統裏的 `WScript.Shell` 對象 ( Object ) 裏的 `.Exec` 方法創建子進程 ( child Process ) 並再調用視窗 ( Windows ) 系統裏的 shell 語句控制臺命令行 ( cmd.exe ) 執行 Bash 語句運行程式設計語言 ( Python ) 解釋器 ( Interpreter ) 的二進位可執行檔 ( .exe ) 加載執行自定義創建的 http 伺服器 ( Server ) 應用代碼脚本檔 ( .py ) 從而實現, 單擊 ( Click ) 微軟電子表格 Microsoft Excel 應用軟體的菜單欄 ( Menu bar ) 裏自定義的子菜單, 即可一鍵快捷啓動使用計算機程式設計語言 ( Python ) 自定義創建的 http 伺服器 ( Server ) 應用的效果.

![]()

---

Operating System :

Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30

---

Application :

Microsoft Office Excel Professional 2019 x86_64

Interpreter :

Node.js - version 22.20.0

npm - version 10.7.0

Database :

MongoDB mongod - version 8.2.3

MongoDB mongosh - version 2.6.0

Interpreter :

Python - version 3.11.2

pip - version 22.3.1

Database :

MariaDB - version 10.11

---

Application : Microsoft Office Excel Professional 2019

[作業系統 ( Operating system ) 之 Microsoft Windows 官方網站](https://www.microsoft.com/zh-tw/windows): 
https://www.microsoft.com/zh-tw/windows

[電子表格應用 Microsoft Office Excel 官方下載頁](https://www.microsoft.com/zh-tw/download/office): 
https://www.microsoft.com/zh-tw/download/office

[電子表格應用 Microsoft Office Excel 2019 官方説明頁](https://learn.microsoft.com/zh-tw/deployoffice/office2019/overview): 
https://learn.microsoft.com/zh-tw/deployoffice/office2019/overview

微軟電子表格 ( Windows - Office - Excel - Visual Basic for Applications ) 應用，轉換 JSON 字符串類型的變量 ( JSON - String Object ) 與微軟電子表格字典類型的變量 ( Windows - Office - Excel - Visual Basic for Applications - Dict Object ) 數據類型，通過借用微軟電子表格 ( Windows - Office - Excel - Visual Basic for Applications ) 應用的第三方擴展類模組「VBA-JSON」實現.

[相互轉換 JSON 字符串與 Excel-VBA-Dict 對象 ( Object ) 使用的第三方擴展類模組 VBA-JSON 官方 GitHub 網站倉庫](https://github.com/VBA-tools/VBA-JSON): 
https://github.com/VBA-tools/VBA-JSON.git

微軟電子表格 ( Windows - Office - Excel - Visual Basic for Applications ) 應用，操作 Windows - Edge , Google - Chrome , Mozilla - Firebox 瀏覽器, 通過借用 codeproject 網站提供的，微軟電子表格 ( Windows - Office - Excel - Visual Basic for Applications ) 應用的第三方擴展類模組 : clsBrowser.cls , clsCore.cls , clsJsConverter.cls 實現，官方網站 ( codeproject ) 的地址 ( Uniform Resource Locator , URL ) 如下 :

[第三方擴展類模組提供網站 codeproject 裏的 Automate Chrome or Edge using VBA 庫 ( Tips ) 官方説明頁](https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA): 
https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA

[第三方擴展類模組 Chromium-Automation-with-CDP-for-VBA 官方 GitHub 網站倉庫](https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA): 
https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA.git

[第三方擴展類模組 Edge-IE-Mode-Automation-with-IES-for-VBA 官方 GitHub 網站倉庫](https://github.com/longvh211/Edge-IE-Mode-Automation-with-IES-for-VBA): 
https://github.com/longvh211/Edge-IE-Mode-Automation-with-IES-for-VBA.git

Interpreter : Node.js

[程式設計 JavaScript 語言解釋器 ( Interpreter ) 之 Node.js 官方網站](https://node.js.org/): 
https://node.js.org/

[程式設計 JavaScript 語言解釋器 ( Interpreter ) 之 Node.js 官方網站](https://nodejs.org/en/): 
https://nodejs.org/en/

[程式設計 JavaScript 語言解釋器 ( Interpreter ) 之 Node.js 官方下載頁](https://nodejs.org/en/download/package-manager): 
https://nodejs.org/en/download/package-manager

[程式設計 JavaScript 語言解釋器 ( Interpreter ) 之 Node.js 官方 GitHub 網站賬戶](https://github.com/nodejs): 
https://github.com/nodejs

[程式設計 JavaScript 語言解釋器 ( Interpreter ) 之 Node.js 官方 GitHub 網站倉庫](https://github.com/nodejs/node): 
https://github.com/nodejs/node.git

Interpreter : Python

[程式設計 Python 語言解釋器 ( Interpreter ) 官方網站](https://www.python.org/): 
https://www.python.org/

[程式設計 Python 語言解釋器 ( Interpreter ) 官方下載頁](https://www.python.org/downloads/): 
https://www.python.org/downloads/

[程式設計 Python 語言解釋器 ( Interpreter ) 官方 GitHub 網站賬戶](https://github.com/python): 
https://github.com/python

[程式設計 Python 語言解釋器 ( Interpreter ) 官方 GitHub 網站倉庫頁](https://github.com/python/cpython): 
https://github.com/python/cpython.git

Database : Microsoft Access

[資料庫 Microsoft Access 應用軟體官方網站](https://www.microsoft.com/en-us/microsoft-365/access): 
https://www.microsoft.com/en-us/microsoft-365/access

[資料庫 Microsoft Access 應用軟體官方網站中文版](https://www.microsoft.com/zh-tw/microsoft-365/access): 
https://www.microsoft.com/zh-tw/microsoft-365/access

[資料庫 Microsoft Access 應用軟體官方手冊](https://learn.microsoft.com/en-us/office/vba/api/overview/access): 
https://learn.microsoft.com/en-us/office/vba/api/overview/access

[資料庫 Microsoft Access 應用軟體官方手冊中文版](https://learn.microsoft.com/zh-tw/office/vba/api/overview/access): 
https://learn.microsoft.com/zh-tw/office/vba/api/overview/access

Database : MongoDB

[資料庫 MongoDB 應用軟體官方網站](https://www.mongodb.com/): 
https://www.mongodb.com/

[資料庫 MongoDB 應用軟體官方手冊](https://www.mongodb.com/docs/manual/): 
https://www.mongodb.com/docs/manual/

[資料庫 MongoDB 應用軟體下載官方網站](https://www.mongodb.com/try/download/community): 
https://www.mongodb.com/try/download/community

[資料庫 MongoDB 應用軟體官方 GitHub 網站賬戶](https://github.com/mongodb): 
https://github.com/mongodb

[資料庫 MongoDB 應用軟體官方 GitHub 網站倉庫](https://github.com/mongodb/mongo): 
https://github.com/mongodb/mongo.git

Database : MariaDB

[資料庫 MariaDB 應用軟體官方網站](https://mariadb.com/): 
https://mariadb.com/

[資料庫 MariaDB 應用軟體官方手冊](https://mariadb.com/docs/server): 
https://mariadb.com/docs/server

[資料庫 MariaDB 應用軟體下載官方網站](https://mariadb.org/download/): 
https://mariadb.org/download/

[資料庫 MariaDB 應用軟體官方 GitHub 網站賬戶](https://github.com/MariaDB): 
https://github.com/MariaDB

[資料庫 MariaDB 應用軟體官方 GitHub 網站倉庫](https://github.com/MariaDB/server): 
https://github.com/MariaDB/server.git

---

開箱即用 ( out of the box ) ( portable application ) 已配置第三方擴展模組 ( third-party extensions ( libraries or modules ) ) 的運行環境的壓縮檔 ( .zip .7z ) 的 [百度網盤(pan.baidu.com)](https://pan.baidu.com/s/1jLLxakrQrE8wpXHlr9GX4w?pwd=yyrf) 下載頁: 
https://pan.baidu.com/s/1jLLxakrQrE8wpXHlr9GX4w?pwd=yyrf

提取碼：yyrf

開箱即用 ( out of the box ) ( portable application ) 檔 :

1. 壓縮檔 : `Nodejs-22.20.0-Window10-AMD_FX8800P_x86_64.7z`

壓縮檔「`Nodejs-22.20.0-Window10-AMD_FX8800P_x86_64.7z`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 程式設計語言 ( JavaScript ) 解釋器 ( Interpreter ) 二進位可執行檔 ( node-v22.20.0-x64.msi ) 開箱即用 ( out of the box ) ( portable application ) 免安裝版，需自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/Nodejs/` 内，最終完整路徑應爲「`Database-to-Excel-VBA/Nodejs/Nodejs-22.20.0/node.exe`」

2. 壓縮檔 : `Python-3.11.2-Window10-AMD_FX8800P_x86_64.7z`

壓縮檔「`Python-3.11.2-Window10-AMD_FX8800P_x86_64.7z`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 程式設計語言 ( Python ) 解釋器 ( Interpreter ) 二進位可執行檔 ( python-3.11.2-amd64.exe ) 開箱即用 ( out of the box ) ( portable application ) 免安裝版，需自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/Python/` 内，最終完整路徑應爲「`Database-to-Excel-VBA/Python/Python311/python.exe`」

3. 壓縮檔 : `NodejsToMongoDB-MongoDB_8.2.3-Window10-AMD_FX8800P_x86_64.zip`

壓縮檔「`NodejsToMongoDB-MongoDB_8.2.3-Window10-AMD_FX8800P_x86_64.zip`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 使用程式設計語言 ( computer programming language ) : JavaScript 鏈接操作 MongoDB 資料庫的伺服器 'NodejsToMongoDB' 開箱即用 ( out of the box ) ( portable application ) 版，已配置計算機程式設計語言 ( computer programming language ) : JavaScript 解釋器 ( Interpreter ) 運行此資料庫伺服器 'NodejsToMongoDB' 項目所需的第三方擴展模組 ( third-party extensions ( libraries or modules ) ) 的運行環境，可自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/MongoDB/NodejsToMongoDB/` 内，再因應協調配置壓縮檔「`Nodejs-22.20.0-Window10-AMD_FX8800P_x86_64.7z`」之後，即可使用如下指令啓動運行資料庫伺服器「`NodejsToMongoDB`」項目 : 
```
C:\Database-to-Excel-VBA\MongoDB> C:/Database-to-Excel-VBA/Nodejs/Nodejs-22.20.0/node.exe C:/Database-to-Excel-VBA/MongoDB/NodejsToMongoDB/Nodejs2MongodbServer.js host=::0 port=27016 number_cluster_Workers=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
```

4. 壓縮檔 : `PythonToMariaDB-MariaDB10.11-Window10-AMD_FX8800P_x86_64.zip`

壓縮檔「`PythonToMariaDB-MariaDB10.11-Window10-AMD_FX8800P_x86_64.zip`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 使用程式設計語言 ( computer programming language ) : Python 鏈接操作 MariaDB 資料庫的伺服器 'PythonToMariaDB' 開箱即用 ( out of the box ) ( portable application ) 版，已配置計算機程式設計語言 ( computer programming language ) : Python 解釋器 ( Interpreter ) 運行此資料庫伺服器 'PythonToMariaDB' 項目所需的第三方擴展模組 ( third-party extensions ( libraries or modules ) ) 的運行環境，可自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/MariaDB/PythonToMariaDB/` 内，再因應協調配置壓縮檔「`Python-3.11.2-Window10-AMD_FX8800P_x86_64.7z`」之後，即可使用如下指令啓動運行統計運算伺服器「'PythonToMariaDB`」項目 : 
```
C:\Database-to-Excel-VBA\MariaDB> C:/Database-to-Excel-VBA/MariaDB/PythonToMariaDB/Scripts/python.exe C:/Database-to-Excel-VBA/MariaDB/PythonToMariaDB/src/Python2MariaDBServer.py host=::0 port=27016 Is_multi_thread=False number_Worker_process=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
```
或者 : 
```
C:\Database-to-Excel-VBA\MariaDB> C:/Database-to-Excel-VBA/Python/Python311/python.exe C:/Database-to-Excel-VBA/MariaDB/PythonToMariaDB/src/Python2MariaDBServer.py host=::0 port=27016 Is_multi_thread=False number_Worker_process=0 MongodbHost=[::1] MongodbPort=27017 dbUser=admin_Database1 dbPass=admin dbName=Database1
```

5. 壓縮檔 : `Server-MongoDB_8.2.3-Window10-AMD_FX8800P_x86_64.zip`

壓縮檔「`Server-MongoDB_8.2.3-Window10-AMD_FX8800P_x86_64.zip`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 資料庫應用 MongoDB 伺服器端二進位可執行啓動檔 'mongod.exe' 開箱即用 ( out of the box ) ( portable application ) 版運行環境，可自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/MongoDB/Server/` 内，最終完整路徑應爲「`Database-to-Excel-VBA/MongoDB/Server/8.2/bin/mongod.exe`」，即可使用如下指令啓動運行資料庫 MongoDB 伺服器應用 : 
```
C:\Database-to-Excel-VBA\MongoDB> C:/Database-to-Excel-VBA/MongoDB/Server/8.2/bin/mongod.exe --config=C:/Database-to-Excel-VBA/MongoDB/NodejsToMongoDB/mongod.cfg
```

6. 壓縮檔 : `mongosh_2.6.0-Window10-AMD_FX8800P_x86_64.zip`

壓縮檔「`mongosh_2.6.0-Window10-AMD_FX8800P_x86_64.zip`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 資料庫應用 MongoDB 用戶端二進位可執行啓動檔 'mongosh.exe' 開箱即用 ( out of the box ) ( portable application ) 版運行環境，可自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/MongoDB/mongosh/` 内，最終完整路徑應爲「`Database-to-Excel-VBA/MongoDB/mongosh/mongosh.exe`」，即可使用如下指令啓動運行資料庫 MongoDB 用戶端應用 : 
```
C:\Database-to-Excel-VBA\MongoDB> C:/Database-to-Excel-VBA/MongoDB/mongosh/mongosh.exe mongodb://username:password@[::1]:27017/Database1
```

7. 壓縮檔 : `data-MongoDB_8.2.3-Window10-AMD_FX8800P_x86_64.zip`

壓縮檔「`data-MongoDB_8.2.3-Window10-AMD_FX8800P_x86_64.zip`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 資料庫應用 MongoDB 伺服器端自定義創建的名爲 'Database1' 資料庫 ( Database ) , 内含名爲 'Collection1' 自定義數據集 ( Collection/Table ) , 開箱即用 ( out of the box ) ( portable application ) 版運行環境，可自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/MongoDB/data/` 内，可使用資料庫 MongoDB 用戶端應用鏈接伺服器之後，操作處理增、刪、改、查資料集合.

8. 壓縮檔 : `MariaDB10.11-Window10-AMD_FX8800P_x86_64.zip`

壓縮檔「`MariaDB10.11-Window10-AMD_FX8800P_x86_64.zip`」爲微軟視窗作業系統 ( Operating System: Acer-NEO-2023 Windows10 x86_64 Inter(R)-Core(TM)-m3-6Y30 ) 資料庫應用 MariaDB 伺服器端二進位可執行啓動檔 'mysqld.exe' 開箱即用 ( out of the box ) ( portable application ) 版運行環境，可自行下載解壓縮，將其保存至檔案夾 ( folder ) : `Database-to-Excel-VBA/MariaDB/MariaDB10.11/` 内，最終完整路徑應爲「`Database-to-Excel-VBA/MariaDB/MariaDB10.11/bin/mysqld.exe`」，即可使用如下指令啓動運行資料庫 MongoDB 伺服器應用 : 
```
C:\Database-to-Excel-VBA\MariaDB> C:/Database-to-Excel-VBA/MariaDB/MariaDB10.11/bin/mysqld.exe
```
即可.
