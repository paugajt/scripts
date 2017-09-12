

intHighNumber = 6
intLowNumber =1
For i = 1 to 1
     Randomize
     intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
     Wscript.Echo intNumber
Next