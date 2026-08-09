# paragraph rendering verification
## normal paragraph
### N1 soft break plain
abc
def

### N2 soft break with emphasis across lines
*foo
bar*

### N3 hard break by two trailing spaces
a*b  
**c**

### N4 hard break by backslash
*foo\
bar*

### N5 explicit br
foo<br>**bar**

### N6 inline code should keep br literal
`aa<br>bb` and **cc**

### N7 inline code around soft break
`aa
bb` and **cc**


## block quote paragraph
### Q1 quote soft break plain
> abc
> def

### Q2 quote soft break with emphasis across lines
> *foo
> bar*

### Q3 quote hard break by two trailing spaces
> a*b  
> **c**

### Q4 quote hard break by backslash
> *foo\
> bar*

### Q5 quote explicit br
> foo<br>**bar**

### Q6 quote blank line separates paragraphs
> first line
>
> second line

### Q7 quote inline code should keep br literal
> `aa<br>bb` and **cc**


## bullet list paragraph
### L1 bullet item soft break plain
- abc
  def

### L2 bullet item soft break with emphasis across lines
- *foo
  bar*

### L3 bullet item hard break by two trailing spaces
- a*b  
  **c**

### L4 bullet item hard break by backslash
- *foo\
  bar*

### L5 bullet item explicit br
- foo<br>**bar**

### L6 bullet item with inline code br literal
- `aa<br>bb` and **cc**

### L7 nested bullet with continuation
- parent
  - child
    tail


## numbered list paragraph
### NU1 numbered item soft break plain
1. abc
  def

### NU2 numbered item hard break by two trailing spaces
1. a*b  
  **c**

### NU3 numbered item hard break by backslash
1. *foo\
  bar*

### NU4 numbered item explicit br
1. foo<br>**bar**

### NU5 numbered item with inline code br literal
1. `aa<br>bb` and **cc**


## quote bullet paragraph
### QB1 quote bullet soft break plain
> - abc
>   def

### QB2 quote bullet hard break by two trailing spaces
> - a*b  
>   **c**

### QB3 quote bullet hard break by backslash
> - *foo\
>   bar*

### QB4 quote bullet explicit br
> - foo<br>**bar**


## quote numbered paragraph
### QN1 quote numbered soft break plain
> 1. abc
>   def

### QN2 quote numbered hard break by two trailing spaces
> 1. a*b  
>   **c**

### QN3 quote numbered hard break by backslash
> 1. *foo\
>   bar*

### QN4 quote numbered explicit br
> 1. foo<br>**bar**

## table rendering
### T1 basic table
| h1 | h2 |
| --- | --- |
| a | b |
| c | d |

### T2 inline formatting in table cells
| h1 | h2 | h3 |
| --- | --- | --- |
| *foo* | **bar** | `baz` |

### T3 explicit br in table cells should expand to next table row
| h1         | h2         |
| ---------- | ---------- |
| foo<br>bar | a<br>**b** |

### T3b table cells with different br counts
| h1          | h2 |
| ----------- | -- |
| a<br>b<br>c | x  |

### T3c mixed br and non-br rows should preserve borders only between markdown rows
| h1     | h2     |
| ------ | ------ |
| plain1 | plain2 |
| a<br>b | x      |
| plain3 | plain4 |
| y      | c<br>d |
| plain5 | plain6 |

### T4 escaped pipe and code pipe
| h1 | h2 |
| --- | --- |
| a \| b | `x|y` |

### T4b escaped pipe inside code span
| h1 | h2 |
| --- | --- |
| a \| b | `x\|y` |

### T5 table boundary with surrounding paragraphs
before table

| h1 | h2 |
| --- | --- |
| a | b |

after table


## code block rendering
### C1 fenced code block basic
```text
line1
line2
```

### C2 code block should not parse markdown
```text
*not italic*
**not bold**
`not code span`
| not | a | table |
- not a list
> not a quote
a*b
**c**
```

### C3 tilde fence
~~~text
alpha
beta
~~~

### C4 code block boundary with surrounding paragraphs
before code

```text
code body
```

after code

### C5 quote and code block boundary
> quoted before code

```text
plain code
```

> quoted after code

### C6 list and code block boundary
- list item before code

```text
code after list
```

normal after code

## mixed block boundaries
### B1 quote followed by table
> quote line before table

| h1 | h2 |
| --- | --- |
| a | b |

### B2 table inside quote should render as table
> | h1 | h2 |
> | --- | --- |
> | a | b |

### B3 quote bullet paragraph followed by table
> - item
>   detail

| h1 | h2 |
| --- | --- |
| x | y |

### B4 table followed by quote
| h1 | h2 |
| --- | --- |
| a | b |

> quote after table

### B5 list followed by table followed by normal paragraph
- list item
  detail

| h1 | h2 |
| --- | --- |
| a | b |

after mixed blocks

### B6 quote then normal then list then table
> quoted start
> second line

normal line after quote

- list after normal
  detail after list

| h1 | h2 |
| --- | --- |
| last | row |