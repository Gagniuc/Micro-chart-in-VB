# Micro chart in VB

This compact chart takes into account both positive and negative values from an input. Thus, this VB chart takes into account a lower bound as well as an upper bound. The lower bound represents the lowest value whereas the upper bound represents the highest value over the input. The project in this repository shows two VB charts and both use the PictureBox object from VB6. The first one found in folder <kbd>src/chart_short</kbd> contains the shortest source code for a chart. Basically the implementation is represented by a function named <kbd>chart</kbd> that draws on a PictureBox object based on some consecutive numeric values. The second chart found in folder <kbd>src/chart</kbd> contains an addition to the first, namely it draws the <kbd>x-axis</kbd> and <kbd>y-axis</kbd>, and the corresponding baseline ticks. For more detailed information, note that these native charts in VB6, were published in the supplementary materials of the book entitled <i>Algorithms in Bioinformatics: Theory and Implementation</i>. The screenshot below shows the output of the function:

![screenshot](https://github.com/Gagniuc/World-shortest-chart-in-VB6/blob/main/img/chart_short.png?raw=true)

The second chart project shows an addition to the first, namely it draws the <kbd>x-axis</kbd> and <kbd>y-axis</kbd>, and the corresponding baseline ticks. Also, the position of the chart can be changed inside the object with the help of four variables responsible for the vertical position, the horizontal position, the width of the chart and the height of the chart. The screenshot below shows the output of the function:

![screenshot](https://github.com/Gagniuc/World-shortest-chart-in-VB6/blob/main/img/chart.png?raw=true)

How does it work? The <kbd>chart</kbd> function contains a loop that makes a number of iterations (<i>i</i>) equal to the number of terms present in the sequence (<i>s</i>). Inside the main loop, the coordinates above the canvas object are calculated based on the maximum value, namely according to the value found in the <i>mx</i> variable. Thus, the <kbd>y-axis</kbd> is represented by the height (<i>h</i>) of the PictureBox object divided by the value in the <i>mx</i> variable (<i>h</i>/<i>mx</i>), and the result is multiplied by the current value in the sequence (s[<i>i</i>]). To position the zero values at the bottom of the chart, the <kbd>y-axis</kbd> is reversed by subtracting the result (the <i>y</i> value) from the height (<i>h</i>) of the PictureBox object. However, for a better visualization, the implementation of this chart narrows the <kbd>y-axis</kbd> and shows only the region between the two values. To obtain this relative reduction, the minimum value was taken into account, namely:

<img src="https://github.com/Gagniuc/World-shortest-chart-in-VB6/blob/main/img/ylu.png?raw=true" height="100">

In contrast, the <kbd>x-axis</kbd> is calculated by dividing the length of the PictureBox object by the total number of terms in the sequence (<b>w/s.length</b>), and the
result is multiplied by the iteration number (<i>i</i>):

<img src="https://github.com/Gagniuc/World-shortest-chart-in-VB6/blob/main/img/x.png?raw=true" height="100">

Where <i>mn</i> is the minimum value and <i>mx</i> is the maximum value found over the input (the consecutive numeric values spaced by delimiters), <i>h</i> is the PictureBox height, and s[<i>i</i>] is the current value from the input. Note that the inner workings of the <kbd>chart</kbd> function were fully described in a [previous JS implementation](https://github.com/Gagniuc/World-smallest-js-chart-v1.0). This concludes the changes related to the <kbd>chart</kbd> function.

```vb6
Function chart(g, c, e)

    sig = Split(g, ",")
    
    mx = 0
    mn = 0
    
    For i = 0 To UBound(sig)
        If (Val(sig(i)) > mx) Then mx = Val(sig(i))
        If (Val(sig(i)) < mn) Then mn = Val(sig(i))
    Next i

    w = graf_val.ScaleWidth
    h = graf_val.ScaleHeight

    d = (w - 80) / (UBound(sig) - 1)
    
    If (e = "|") Then
        graf_val.Cls
        mxg = mx
        mng = mn
    End If
    
    graf_val.DrawWidth = 4
    
    For i = 0 To UBound(sig) - 1
    
        y = h - 15 - ((h - 15) / (mx - mn)) * (Val(sig(i)) - mn)
        x = d * i

        If (i = 0) Then
            oldX = x
            oldY = y
        End If
        
        graf_val.Line (oldX, oldY)-(x, y), c
        
        oldX = x
        oldY = y
        
    Next i
 
End Function
```

The lines below show how this <kbd>Chart</kbd> function from above can be called:

```vb6
A = "0,0.14,0.29,0.45,0.64,0.86,1.14,1.53,2.13,3.27,6.41,75.31,7.75,3.61,2.29,1.62,1.2,0.9,0.67,0.48,0.32,0.17,0.03,0.12,0.26,0.42,0.6,0.81,1.08,1.44,2,2.99,5.45,25.09,9.79,4.03,2.47,1.72,1.27,0.95,0.71,0.52,0.35,0.2,0.05,0.09,0.23,0.39,0.56,0.77,1.02,1.36,1.87,2.74,4.74,15.04,13.27,4.54,2.67,1.83,1.34"

Call chart(A, vbRed, "|")
```

# References

<i>Paul A. Gagniuc. Algorithms in Bioinformatics: Theory and Implementation. John Wiley & Sons, Hoboken, NJ, USA, 2021, ISBN: 9781119697961.</i>
