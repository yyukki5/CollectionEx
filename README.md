# CollectionEx
Expanding Collection's function.


## How to use
~~~
    res1 = CollectionEx(col) _
        .Where("x=>x.abc<7") _
        .OrderByDescending("x=>x.abc") _
        .Take(3) _
        .SelectBy("x=>x.abc") _
        .ToArray
    
    set res2 = CollectionEx.Init(col).Take(3).Items
~~~

## Features
 - No need to write "Dim" and "New" for use.  class is predeclared, and can initialize by class name as default function. 
 - After initialization, run some function and output as Collection or something
 -  Some functions can use "lambda string" (Simply implemented anonymous function by string)
 
## Files
 - Can import these files to your VBA project
    - CollectionEx.cls
    - Lambda.cls (<- see [Lambda Repository](https://github.com/yyukki5/Lambda)  (Ver.0.6.0))
 - For only sample
    - Sample.bas
    - Class1.cls
    - Class2.cls

<br>
 * Specification may be changed until version 1.0.0.

## Japanese Note
Collection を使う時に少しラクするためのクラス  
Collection をよく使うので、少し楽したいなと思ったので自作  
LINQライクに出来るといいかなと思ったが、まだ計算できないときがあるかも...  
→ Lambda.clsは別Repositoryで作ることにしました

<br>
作成中


# ColEx
Simply CollectionEx


## How to use
~~~vb
    Dim col as new Collection
    if ColEx(col).Where(...).Skip(3).Take(3).AnyBy(...) then
        Debug.Pint ColEx(col).SelectBy("").First()
    end if
    
    Dim v as Class1
    For Each v in ColEx(col).Take(3)
        Debug.Print v.Prop1
    Next

    Dim cex as New ColEx
    Call cex.Initialize(col)
~~~

## Features
1. Predicred
1. Quick initialized (Clone())
1. NewEnum
1. Any LINQ functions
    - Where()
    - SelectBy()
    - AnyBy(), AllBy()
    - Take(), Skip()


## Japanese Note
CollectionEx を簡略化、軽量化したもの。  

要素の直下のプロパティを使ってWhereしたり、Selectしたり出来る。  
2階層以降のプロパティにはアクセスできないので For Each を使う。  
デフォルトメソッドはCreate()
