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
        .Where("Abc", cexLessThan, 7) _
        .OrderByDescending("Abc") _
        .Take(3) _
        .SelectBy("Def.Def") _
        .ToArray
    
    Dim v as Class1
    For Each v in ColEx(col).Take(3)
        Debug.Print v.Prop1
    Next

    Call ColEx(Sheet1.Range("A1:A10")).Where("Value", cexEqual, Empty).SelectBy("Delete", VbMethod)
~~~

## Features
1. Predeclared
1. Quick initialized (used default function "Create()")
1. Can reference property of class by using property names. 
1. Can use "For Each"
1. Any LINQ functions
    - Where()
    - SelectBy(), SelectManyBy()
    - AnyBy(), AllBy()
    - Take(), Skip(), First(), FirstOrDafult(), Last(), LastOrDefault(), SingleBy(), SingleOrDefaultBy()  
    - OrderBy(), OrderByDescending()
    - Contains(), Distinct(), DistinctBy()

<br>
 * If comparing objects as equal, objects should has Equals() function. if not has, raise error. <br>
 * Specification may be changed until version 1.0.0.  




## Japanese Note
CollectionEx を簡略化、軽量化したもの。（Lambda を使っていない）  

- 要素のプロパティを使ってWhereしたり、Selectしたり出来る  
- SelectBy() を除き、メソッドを使う場合はForEach を利用するかCollectionEx を使う 
- SelectBy() はメソッドを指定できます。（指定した最下層のクラスのみ有効）
- デフォルトメソッドはCreate()

<br>
作成中