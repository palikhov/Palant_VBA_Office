Function Rolldice(dice, faces, adder)
Randomize
Rolldice = 0

For Count = 1 To dice
Rolldice = Rolldice + Int(Rnd * faces) + 1
Next Count

Rolldice = Rolldice + adder

End Function
