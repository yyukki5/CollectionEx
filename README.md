# CollectionEx
Expanding Collection's function.


## How to use
~~~
    res1 = CollectionEx(col) _
        .Where("x => x.abc < 7") _
        .OrderByDescending("x => x.abc") _
        .Take(3) _
        .SelectBy("x => x.abc") _
        .ToArray
    
    set res2 = CollectionEx.Initialize(col).Take(3).Items
~~~

## Features
 - No need to write "Dim" and "New" for use.  class is predeclared, and can initialize by class name as default function. 
 - After initialization, run some function and output as Collection or something
 - Some functions can use "lambda string" (Simply implemented anonymous function by string) 
 - If comparing objects as equal, objects should has Equals() function. if not has, raise error
 
<br>
 * Specification may be changed until version 1.0.0.  

 
## Files
 - Can import these files to your VBA project
    - CollectionEx.cls
    - Lambda.cls (<- see [Lambda Repository](https://github.com/yyukki5/Lambda)  (Ver.0.6.0))




## Japanese Note
Collection を使う時に少しラクするためのクラス  
Collection をよく使うので、少し楽したいなと思ったので自作  
LINQライクに出来るといいかなと思ったが、まだ計算できないときがあるかも...  
→ Lambda.clsは別Repositoryで作ることにしました


  
作成中


# ColEx
Simplified and Faster CollectionEx


## How to use
~~~vb
    Dim res
    res = ColEx(col) _
        .Where("abc", cexLessThan, 7) _
        .OrderByDescending("abc") _
        .Take(3) _
        .SelectBy("abc") _
        .ToArray
    
    Dim v as Class1
    For Each v in ColEx(col).Take(3)
        Debug.Print v.Prop1
    Next

    Dim cex as New ColEx
    Call cex.Initialize(col)
~~~

## Features
1. Predeclared
1. Quick initialized (used default function "Create()")
1. can use "For Each"
1. Any LINQ functions
    - Where()
    - SelectBy(), SelectManyBy()
    - AnyBy(), AllBy()
    - Take(), Skip()
    - OrderBy(), OrderByDescending()
    - Contains(), Distinct(), DistinctBy()

<br>
 * If comparing objects as equal, objects should has Equals() function. if not has, raise error. <br>
 * Specification may be changed until version 1.0.0.  




## Japanese Note
CollectionEx を簡略化、軽量化したもの。（Lambda を使っていない）  

- 要素の直下のプロパティを使ってWhereしたり、Selectしたり出来る  
- 2階層以降のプロパティにはアクセスできないので For Each を使う  
- メソッドを使う場合は、ForEachを利用するか、CollectionExを使う
- デフォルトメソッドはCreate()

<br>
作成中