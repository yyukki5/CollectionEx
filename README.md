# CollectionEx
Expanding Collection's function.


## How to use
~~~
    res1 = CollectionEx.Init(col) _
        .WhereByEvaluatedLambda("x=>x.abc<7") _
        .OrderByDescending("x=>x.abc") _
        .Take(3) _
        .SelectByLambda("x=>x.abc") _
        .ToArray
    
    set res2 = CollectionEx.Init(col).Take(3).Items
~~~

 - No need to "Dim" and "New" for use (Predeclared) 
 - After initialization, run some function and output as Collection or something
 -  Some functions can use tiny lambda & delegate string
 
## Files
 - Can import your project
    - CollectionEx.cls
    - Lambda.cls (<- maybe delete in future)
 - For only sample
    - Sample.bas
    - Class1.cls
    - Class2.cls

<br>
 * Specification may be changed.

## Japanese Note
Collection を使う時に少しラクするためのクラス  
最近コレクションをよく使うので、少し楽したいなと思ったので自作  
LINQライクに出来るといいかなと思ったが、まだ計算できないときがあるかも...

<br>
作成中
