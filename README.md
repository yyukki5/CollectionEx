# CollectionEx
Expanding Collection's function.


## How to use
~~~vb
    res1 = CollectionEx(col) _
        .Where("x => x.abc < 7") _
        .OrderByDescending("x => x.abc") _
        .Take(3) _
        .SelectBy("x => x.abc") _
        .ToArray()
    
    set res2 = CollectionEx.Initialize(col).Take(3).Items
~~~

## Features
- Predeclared & Quick initialized. <br> No need to write "Dim" and "New" for use.  class is predeclared, and can initialize by class name as default function. 
- Can use "For Each"
- Can use some LINQ likes functions
- Can refer property or method by using "lambda string" (Simply implemented anonymous function by string. example: ```"x => x.abc"``` ) 

<br>
 * If comparing objects as equal, objects should has Equals() function. if not has, raise error. <br>


## Files
1. CollectionEx.cls
1. Lambda.cls (<- see [Lambda Repository](https://github.com/yyukki5/Lambda) )

After Importing these 2 files to your VBA project, you can use it.



## Japanese Note
- Collection を使う時に少しラクするためのクラス  
- 宣言済み かつ デフォルトメソッドが Create()なので、短い記述で使えます
- Collection と同じように For Each でループを回すことも出来ます
- ラムダ文字列を用いて、LINQのようなメソッドが使えます
- CollectionEx と Lambda のファイルを２つインポートするだけで使えます 


<br>

# ColEx
More Simply and Faster than CollectionEx


## How to use
~~~vb
    Dim res
    res = ColEx(col) _
        .Where("Abc", cexLessThan, 7) _
        .OrderByDescending("Abc") _
        .Take(3) _
        .SelectBy("Def.Def") _
        .ToArray()
    
    Dim v as Class1
    For Each v in ColEx(col).Take(3)
        Debug.Print v.Prop1
    Next

    Call ColEx(Sheet1.Range("A1:A10")).Where("Value", cexEqual, Empty).SelectBy("Delete", VbMethod)
~~~

## Features
- Predeclared & Quick initialized (used default function "Create()")
- Can use "For Each"
- Can use some LINQ likes functions
- Can refer property or method of class by using property names. example: ```"Abc"``` 
### LINQ likes functions
- Where()
- SelectBy(), SelectManyBy()
- AnyBy(), AllBy()
- First(), FirstOrDefault(), Last(), LastOrDefault(), SingleBy(), SingleOrDefaultBy()
- Take(), Skip()
- OrderBy(), OrderByDescending()
- Contains(), Distinct(), DistinctBy()
- Min(), MinBy(), Max(), MaxBy(), Sum() 

<br>
 * If comparing objects as equal, objects should has Equals() function. if not has, raise error. <br>


 ## Files
1. ColEx.cls

After Importing only 1 file to your VBA project, you can use it.


## Japanese Note
CollectionEx を簡略化、軽量化したもの  
- 宣言済み かつ デフォルトメソッドが Create()なので、短い記述で使えます
- Collection と同じように For Each でループを回すことも出来ます
- 要素のプロパティを使ってLINQのようなメソッド(Where, Selectなど)が使えます  
    - SelectBy() はメソッドを指定できます。（指定した最下層のクラスのみ有効）
    - SelectBy() を除き、メソッドを使う場合はForEach を利用するかCollectionEx を使う 
- ColEx.cls のファイルを１つインポートするだけで使えます 