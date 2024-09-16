/************************************************************************
 * @description ADODB操作类
 * @author 
 * @date 2024/09/13
 * @version 0.0.0
 ***********************************************************************/
Class ADODB {
	static ConnectionString(filePath, HDR := true, readOnly := false, password := "", &fileType := "")
	{
		HDR := HDR ? "Yes" : "No"
		SplitPath filePath, , , &ext
		Switch ext
		{
		Case "xls":
			fileType := "excel"
			if readOnly
				connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Extended Properties="Excel 8.0;HDR=' HDR ';IMEX=1";'    ;只读
			else
				connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Extended Properties="Excel 8.0;HDR=' HDR '";'
		Case "xlsx":
			fileType := "excel"
			if readOnly
				connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Extended Properties="Excel 12.0 Xml;HDR=' HDR ';IMEX=1";'    ;只读
			else
				connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Extended Properties="Excel 12.0 Xml;HDR=' HDR '";'
		Case "xlsb":
			fileType := "excel"
			connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Extended Properties="Excel 12.0;HDR=' HDR '";'
		Case "xlsm":
			fileType := "excel"
			connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Extended Properties="Excel 12.0 Macro;HDR=' HDR '";'
		Case "mdb","accdb":
			fileType := "access"
			if password
				connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Jet OLEDB:Database Password=' password ';'
			else
				connStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' filePath ';Persist Security Info=False'
		Default:
			return
		}
		return connStr
	}

	;Connection Open
	;HDR=Yes 表示电子表格含有标题行，如果没有标题行则设置 HDR 为 No。
	;参考www.connectionstrings.com
	static Open(filePath, HDR := true, readOnly := false, password := "") {
		try	{
			conn := ComObject("ADODB.Connection")
			conn.Open(ADODB.ConnectionString(filePath, HDR, readOnly, password, &fileType))
		} catch
			return
		return ADODB(filePath, fileType, conn, readOnly)
	}

	;创建(隐式)
	__New(filePath, fileType, conn, readOnly) {
		this.filePath := filePath
		this.fileType := fileType
		this.conn := conn
		this.readOnly := readOnly
		this.tableNames := this.UpdateTableNames()
		return this
	}

	;销毁
	__Delete() => this.Quit()
	
	;退出
	Quit() {
		if this.conn.State
			this.conn.Close()
	}

	;打开记录集
	Recordset(source, cursorType := 3, lockType := 3, options?) {
		source := this.tableNames.Has(source) ? ("Select * FROM [" source . (this.fileType = "excel" ? "$]" : "]")) : source
		try {
			rs := ComObject("ADODB.Recordset")
			rs.Open(source, this.conn, this.readOnly ? 3 : cursorType, this.readOnly ? 1 : lockType, options?)
		} catch
			return
		return rs
	}


	;指定 OpenSchema 方法所检索的架构 Recordset 的类型
	;参考 https://docs.microsoft.com/zh-cn/office/client-developer/access/desktop-database-reference/openschema-method-ado
	OpenSchema(options*) {
		try rsSchema := this.conn.OpenSchema(options*)
		catch
			return
		return rsSchema
	}

	;获取所有的表的名称
	UpdateTableNames() {
		this.tableNames := Map()
		try rsSchema := this.conn.OpenSchema(20)
		catch
			return this.tableNames
		while !rsSchema.EOF {
			if (rsSchema.Fields.Item("TABLE_TYPE").Value = "TABLE")
			and (tableName := rsSchema.Fields.Item("TABLE_NAME").Value)
			and !InStr(tableName, "FilterDatabase") {    ;排除筛选表
				tableName := (this.fileType = "excel") ? SubStr(tableName, 1, StrLen(tableName)-1) : tableName
				this.tableNames[tableName] := true
			}
			rsSchema.MoveNext
		}
		rsSchema.Close()
		return this.tableNames
	}

	;读取数据库内容
	static Read(filePath, HDR := false, password := "") {
		;打开文件
		try {
			conn := ComObject("ADODB.Connection")
			conn.Open(ADODB.ConnectionString(filePath, HDR, true, password, &fileType))
		} catch
			return
		;获取表信息
		db := Map()
		try rsSchema := conn.OpenSchema(20)    ;TABLE_CATALOG TABLE_SCHEMA TABLE_NAME TABLE_TYPE
		catch
			return
		while !rsSchema.EOF {
			if (rsSchema.Fields("TABLE_TYPE").Value = "TABLE")
			and (tableName := rsSchema.Fields("TABLE_NAME").Value)
			and !InStr(tableName, "FilterDatabase") {    ;排除筛选表
				tableName := (fileType = "excel") ? SubStr(tableName, 1, StrLen(tableName)-1) : tableName
				db[tableName] := Map()
			}
			rsSchema.MoveNext
		}
		rsSchema.Close()
		;向各表添加信息
		for tableName, table in db {
			source := (fileType = "excel") ? tableName "$" : tableName
			try {
				rs := ComObject("ADODB.Recordset")
				rs.Open("Select * FROM [" tableName . (fileType = "excel" ? "$]" : "]"), conn)
				while !rs.EOF {
					row := table[A_index] := Map()
					loop rs.Fields.Count {       ;获取字段名(从0开始)
						fieldName := rs.Fields(A_Index-1).Name
						row[fieldName] := String(rs.Fields.Item(fieldName).Value)
						;MsgBox rs.Fields(A_Index-1).Name "`n" rs.Fields(A_Index-1).Type
					}
					rs.MoveNext
				}
				rs.Close()
			} catch
				continue
		}
		conn.Close()
		return db
	}

	;写入内容到数据库
	static Write(db, filePath, HDR := false, password := "") {
		;打开文件
		try {
			conn := ComObject("ADODB.Connection")
			conn.Open(ADODB.ConnectionString(filePath, HDR, false, password, &fileType))
		}
		catch
			return
		;向各表添加信息
		for tableName, table in db {
			tableName := (fileType = "excel") ? "[" tableName "$]" : "[" tableName "]"

			rs := ComObject("ADODB.Recordset")
			rs.Open("Select * FROM " tableName, conn, 0, 3)
			while !rs.EOF {
				if table.Has(A_index) {
					for fieldName, value in table[A_index]
						try rs.Fields.Item(fieldName).Value := value
				}
				rs.MoveNext
			}
			rs.Close()
		}
		conn.Close()
		return true
	}
}



/*

解决ADO读取Excel，数据丢失、数据错误、数据乱码问题:
https_blog.csdn.net/?url=https%3A%2F%2Fblog.csdn.net%2Ffranktan2010%2Farticle%2Fdetails%2F28602071

修改 → TypeGuessRow=0

 Excel 97
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\3.5\Engines\Excel
 Excel 2000 and later versions
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel

HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Access Connectivity Engine\Engines\Excel




;https://www.cnblogs.com/hayashi/p/9061498.html
;https://www.microsoft.com/zh-CN/download/details.aspx?id=13255
若要使用此下载，请执行以下操作：
如果您是应用程序用户，请查阅您的应用程序文档，了解有关如何使用相应驱动程序的详细信息。
如果您是使用 OLEDB 的应用程序开发人员，请将 connectionString 属性的 Provider 参数设置为“Microsoft.ACE.OLEDB.12.0”。

如果要连接到 Microsoft Office Excel 数据，请根据 Excel 文件类型添加相应的 OLEDB 连接字符串扩展属性：

文件类型（扩展名）                                             扩展属性
---------------------------------------------------------------------------------------------
Excel 97-2003 工作簿 (.xls)                                  “Excel 8.0”
Excel 2007-2010 工作簿 (.xlsx)                               “Excel 12.0 Xml”
启用宏的 Excel 2007-2010 工作簿 (.xlsm)                      “Excel 12.0 宏”
Excel 2007-2010 非 XML 二进制工作簿 (.xlsb)                  “Excel 12.0”

如果您是使用 ODBC 连接到 Microsoft Office Access 数据的应用程序开发人员，请将连接字符串设置为“Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=path to mdb/accdb file”
如果您是使用 ODBC 连接到 Microsoft Office Excel 数据的应用程序开发人员，请将连接字符串设置为“Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=path to xls/xlsx/xlsm/xlsb file”
*/


/*
connection.Open connectionString,userID,password,options
参数	                描述
connectionString	可选。一个包含有关连接的信息的字符串值。该字符串由一系列被分号隔开的 parameter=value 语句组成的。
userID	            可选。一个字符串值，包含建立连接时要使用的用户名称。
password	        可选。一个字符串值，包含建立连接时要使用的密码。
options	            可选。一个 ConnectOptionEnum 值，确定应在建立连接之后（同步）还是应在建立连接之前（异步）返回本方法。

connectionString
参数	                描述
Provider	        用于连接的提供者的名称。
File Name	        提供者特有的文件（例如，持久保留的数据源对象）的名称，这些文件中包含预置的连接信息。
Remote Provider	    当打开客户端连接时使用的提供者的名称。（仅限于远程数据服务。）
Remote Server	    当打开客户端连接时使用的服务器的路径名称。（仅限于远程数据服务。）
url	                标识资源（比如文件或目录）的绝对 URL。

ConnectOptionEnum
常数	                    值	    描述
adConnectUnspecified	-1	    默认值。同步打开连接。
adAsyncConnect	        16	    异步打开连接。ConnectComplete 事件可以用来确定连接何时可用。
*/

	/*
	;https://learn.microsoft.com/en-us/sql/ado/reference/ado-api/open-method-ado-recordset?view=sql-server-ver16

	recordset.Open Source, ActiveConnection, CursorType, LockType, Options

	* Source
	Optional. 一个变量，其计算结果为一个有效的命令对象、SQL语句、表名、存储过程调用、URL或包含持久存储的记录集的文件或流对象的名称。

	* ActiveConnection
	Optional. 要么是一个有效Connection对象变量名的Variant，要么是一个包含ConnectionString参数的String。

	* CursorType
	Optional. 一个CursorTypeEnum值，该值确定提供程序在打开记录集时应该使用的游标类型。默认值为adOpenForwardOnly。
		adOpenDynamic			2	使用动态游标。其他用户的添加、更改和删除是可见的，并且允许对记录集进行所有类型的移动，如果提供者不支持书签，则书签除外。
		adOpenForwardOnly		0	默认. 使用仅向前游标。与静态游标相同，只是您只能向前滚动记录。当您只需要遍历一个记录集一次时，这将提高性能。
		adOpenKeyset			1	使用键集游标。类似于动态游标，只不过您不能看到其他用户添加的记录，尽管其他用户删除的记录无法从您的记录集中访问。其他用户的数据更改仍然可见。
		adOpenStatic			3	使用静态游标，这是一组记录的静态副本，可用于查找数据或生成报告。其他用户的添加、更改或删除是不可见的。
		adOpenUnspecified  	   -1	不指定游标的类型.

	* LockType
	Optional. 一个LockTypeEnum值，用于确定提供者在打开记录集时应该使用哪种类型的锁定(并发)。缺省值为adLockReadOnly。
		adLockBatchOptimistic	4	表示乐观批处理更新。需要批量更新模式。
		adLockOptimistic		3	指示乐观锁定，逐个记录。提供程序使用乐观锁定，仅在调用Update方法时锁定记录。
		adLockPessimistic		2	表示悲观锁定，逐条记录。提供程序执行确保成功编辑记录所需的操作，通常是在编辑后立即将记录锁定在数据源上。
		adLockReadOnly			1	默认. 只读记录。您不能更改数据。
		adLockUnspecified	   -1	没有指定锁的类型。对于克隆，创建的克隆具有与原始的相同的锁类型。

	* Options
	Optional. 一个长值，指示提供者应该如何计算Source参数(如果它表示Command对象以外的东西)，或者应该从以前保存记录集的文件中恢复记录集。可以是一个或多个CommandTypeEnum或ExecuteOptionEnum值，可以与按位的或操作符组合。
		CommandTypeEnum
			adCmdUnspecified		   -1		不指定命令类型参数。
			adCmdText					1		将CommandText计算为命令或存储过程调用的文本定义。
			adCmdTable					2		将CommandText计算为一个表名，其列全部由内部生成的SQL查询返回。
			adCmdStoredProc				4		将CommandText计算为存储过程名。
			adCmdUnknown				8		默认. 指示CommandText属性中的命令类型未知。
												当命令类型未知时，ADO将多次尝试解释CommandText。
													- CommandText被解释为命令或存储过程调用的文本定义。这是与adCmdText相同的行为。
													- CommandText是存储过程的名称。这与adCmdStoredProc的行为相同。
													- CommandText被解释为表的名称。所有列都由内部生成的SQL查询返回。这与adCmdTable的行为相同。
			adCmdFile					256		将CommandText计算为持久存储的记录集的文件名。与记录集一起使用。只打开或请求。
			adCmdTableDirect			512		将CommandText计算为返回所有列的表名。与记录集一起使用。只打开或请求。要使用Seek方法，必须使用adCmdTableDirect打开记录集。
												该值不能与ExecuteOptionEnum值adAsyncExecute组合使用。

		ExecuteOptionEnum
			adAsyncExecute				16		指示命令应异步执行。
												该值不能与ExecuteOptionEnum值adAsyncExecute组合使用。
			adAsyncFetch				32		指示应异步检索在CacheSize属性中指定的初始数量之后的其余行。
			adAsyncFetchNonBlocking		64		主线程检索时从不阻塞。如果没有检索到请求的行，则当前行自动移动到文件的末尾。
												如果你从一个包含持久存储的记录集的流中打开一个记录集，adAsyncFetchNonBlocking将没有影响;操作将是同步的和阻塞的。
												当使用adCmdTableDirect选项打开记录集时，adAsynchFetchNonBlocking没有作用。
			adExecuteNoRecords			128		命令文本是不返回行的命令或存储过程(例如，只插入数据的命令)。如果检索到任何行，它们将被丢弃而不返回。
												adExecuteNoRecords只能作为可选参数传递给命令或连接执行方法。
			adExecuteStream				1024	以流的形式返回命令执行结果。
												adExecuteStream只能作为可选参数传递给Command Execute方法。
			adExecuteRecord				2048	表示CommandText是一个命令或存储过程，返回一行，应该作为Record对象返回。
			adOptionUnspecified        -1		未指定命令。
	*/

	/*
	Recordset属性:
		AbsolutePage		设置或返回一个可指定 Recordset 对象中页码的值。
		AbsolutePosition	设置或返回一个值，此值可指定 Recordset 对象中当前记录的顺序位置（序号位置）。
		ActiveCommand		返回与 Recordset 对象相关联的 Command 对象。
		ActiveConnection	如果连接被关闭，设置或返回连接的定义，如果连接打开，设置或返回当前的 Connection 对象。
		BOF					如果当前的记录位置在第一条记录之前，则返回 true，否则返回 fasle。
		Bookmark			设置或返回一个书签。此书签保存当前记录的位置。
		CacheSize			设置或返回能够被缓存的记录的数目。
		CursorLocation		设置或返回游标服务的位置。
		CursorType			设置或返回一个 Recordset 对象的游标类型。
		DataMember			设置或返回要从 DataSource 属性所引用的对象中检索的数据成员的名称。
		DataSource			指定一个包含要被表示为 Recordset 对象的数据的对象。
		EditMode			返回当前记录的编辑状态。
		EOF					如果当前记录的位置在最后的记录之后，则返回 true，否则返回 fasle。
		Filter				返回一个针对 Recordset 对象中数据的过滤器。
		Index				设置或返回 Recordset 对象的当前索引的名称。
		LockType			设置或返回当编辑 Recordset 中的一条记录时，可指定锁定类型的值。
		MarshalOptions		设置或返回一个值，此值指定哪些记录被返回服务器。
		MaxRecords			设置或返回从一个查询返回 Recordset 对象的的最大记录数目。
		PageCount			返回一个 Recordset 对象中的数据页数。
		PageSize			设置或返回 Recordset 对象的一个单一页面上所允许的最大记录数。
		RecordCount			返回一个 Recordset 对象中的记录数目。
		Sort				设置或返回一个或多个作为 Recordset 排序基准的字段名。
		Source				设置一个字符串值，或一个 Command 对象引用，或返回一个字符串值，此值可指示 Recordset 对象的数据源。
		State				返回一个值，此值可描述是否 Recordset 对象是打开、关闭、正在连接、正在执行或正在取回数据。
		Status				返回有关批更新或其他大量操作的当前记录的状态。
		StayInSync			设置或返回当父记录位置改变时对子记录的引用是否改变。

	Recordset方法:
		AddNew				创建一条新记录。
		Cancel				撤销一次执行。
		CancelBatch			撤销一次批更新。
		CancelUpdate		撤销对 Recordset 对象的一条记录所做的更改。
		Clone				创建一个已有 Recordset 的副本。
		Close				关闭一个 Recordset。
		CompareBookmarks	比较两个书签。
		Delete				删除一条记录或一组记录。
		Find				搜索一个 Recordset 中满足指定某个条件的一条记录。
		GetRows				把多条记录从一个 Recordset 对象中拷贝到一个二维数组中。
		GetString			将 Recordset 作为字符串返回。
		Move				在 Recordset 对象中移动记录指针。
		MoveFirst			把记录指针移动到第一条记录。
		MoveLast			把记录指针移动到最后一条记录。
		MoveNext			把记录指针移动到下一条记录。
		MovePrevious		把记录指针移动到上一条记录。
		NextRecordset		通过执行一系列命令清除当前 Recordset 对象并返回下一个 Recordset。
		Open				打开一个数据库元素，此元素可提供对表的记录、查询的结果或保存的 Recordset 的访问。
		Requery				通过重新执行对象所基于的查询来更新 Recordset 对象中的数据。
		Resync				从原始数据库刷新当前 Recordset 中的数据。
		Save				把 Recordset 对象保存到 file 或 Stream 对象中。
		Seek				搜索 Recordset 的索引以快速定位与指定的值相匹配的行，并使其成为当前行。
		Supports			返回一个布尔值，此值可定义 Recordset 对象是否支持特定类型的功能。
		Update				保存所有对 Recordset 对象中的一条单一记录所做的更改。
		UpdateBatch			把所有 Recordset 中的更改存入数据库。请在批更新模式中使用。
	*/