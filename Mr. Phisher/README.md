# Mr. Phisher
Easy subscriber TryHackMe room about phishing and malware in office word
https://tryhackme.com/room/mrphisher

Writeup by KyootyBella

## Challenge

"I received a suspicious email with a very weird-looking attachment. It keeps on asking me to "enable macros". What are those?"

A colleague has downloaded a document from a phishing email, but it's telling them to enable macros, from those 2 words, we'll know it's an office program
  

## Enumeration

Whenever we open the room we get access to a word document and a zip file, inside the zip file we'll see the same document.

Opening up the word document we'll see that we are told "this document contains macros"
![Macro_Warning.jpg](https://github.com/KyootyBella/THM-Writeups/blob/main/Mr.%20Phisher/Macro_Warning.jpg)

Hmmm, maybe we should look at the macros?
To locate macros in word documents:
>Tools -> Macros -> Edit Macros

here we'll locate the macro for the document, we find it here
![Macro_Location](https://github.com/KyootyBella/THM-Writeups/blob/main/Mr.%20Phisher/Macro_Location.jpg)
>MrPhisher.docm -> Project -> Modules -> NewMacros

Here's the code by itself
```vb
Option VBASupport 1
Sub Format()
	Dim a()
	Dim b As String
	a = Array(102, 109, 99, 100, 127, 100, 53, 62, 105, 57, 61, 106, 62, 62, 55, 110, 113, 114, 118, 39, 36, 118, 47, 35, 32, 125, 34, 46, 46, 124, 43, 124, 25, 71, 26, 71, 21, 88)
	For i = 0 To UBound(a)
		b = b & Chr(a(i) Xor i)
	Next
End Sub
```

## Time to uncover
Now that we have our malicious code, we need to uncover the hidden message

### The intended way
Let's read through the code and see what it does

```vb
a = Array(102, 109, 99, 100, 127, 100, 53, 62, 105, 57, 61, 106, 62, 62, 55, 110, 113, 114, 118, 39, 36, 118, 47, 35, 32, 125, 34, 46, 46, 124, 43, 124, 25, 71, 26, 71, 21, 88)
For i = 0 To UBound(a)
	b = b & Chr(a(i) Xor i)
```
hmm, looks like we have an array that gets run through for each number that gets xor'ed and converted into ascii characters

Let's see if we can reverse engineer it, I choose to write my solve script in python
so to convert it, it'll look like this
```python
a = [102, 109, 99, 100, 127, 100, 53, 62, 105, 57, 61, 106, 62, 62, 55, 110, 113, 114, 118, 39, 36, 118, 47, 35, 32, 125, 34, 46, 46, 124, 43, 124, 25, 71, 26, 71, 21, 88] 
b = "" 
for i in range(len(a)): 
	b += chr(a[i] ^ i)
```

now that'll go through and write all the ascii characters in our b list, and if we just add a print(b) after the for loop we'll have our flag

```python
a = [102, 109, 99, 100, 127, 100, 53, 62, 105, 57, 61, 106, 62, 62, 55, 110, 113, 114, 118, 39, 36, 118, 47, 35, 32, 125, 34, 46, 46, 124, 43, 124, 25, 71, 26, 71, 21, 88] 
b = "" 
for i in range(len(a)): 
	b += chr(a[i] ^ i)
print(b)
```
(if you want a more cool looking way of getting the flag, add the print statement in the for loop and the flag will gradually get printed)

when we now run the script we'll get flag
![Flag_Script.JPG](https://github.com/KyootyBella/THM-Writeups/blob/main/Mr.%20Phisher/Flag_script.jpg)

EYYYY, WE ACTUALLY GOT FLAG!
`flag{redacted}`


### The interesting way
#### DO ONLY DO THIS WITH CODE YOU KNOW OR HAVE IN AN ISOLATED ENVIRONMENT
After not wanting to debug code and write my own, I decided to do dynamic analysis, aka add a breakpoint and make code go brrrr

The security on the word document has made it so we can't run the macro. so we'll have to change the security settings.
>Tools -> Options -> Security -> macro security

![Security.jpg](https://github.com/KyootyBella/THM-Writeups/blob/main/Mr.%20Phisher/Security.jpg)
when we have changed the security level to low we'll apply and close down the word document and open it again.

Now we can run the macro, but first, let's put a breakpoint in the script so it doesn't end.
Plus follow the "b" value

To add a breakpoint, set your point on the line you want to break at and press F9 or the red circle with a cursor over.
To watch/follow the b value hower over b and press F7 or press the eye.

Macro editor should look like this
![Finished_Script.jpg](https://github.com/KyootyBella/THM-Writeups/blob/main/Mr.%20Phisher/Finished_Script.jpg)
We have the b in watch down in the left corner and a red circle on line 10

We will now run our script and get this output.
![Flag.jpg](https://github.com/KyootyBella/THM-Writeups/blob/main/Mr.%20Phisher/Flag.jpg)

WE GOT FLAG!!
`flag{redacted}`
