
    ucAutoResize.ctl


    Based on the idea of the very good submission from Hamed Oveisi (thx man!)
    (look/vote at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=49740&lngWId=1)
    I tried to improve/change some things ('Refactoring  is the trendy word for ;)  ):


    *  More possibilities in resizing (look to the "3D" lines (frames) ).
    *  No more flickering when forms gets too small.
    *  No more call in the Form_Resize event neccessary. (In fact you didn't need ANY code.)
       (Ups! Just check out Hamed's last update. He did the same ;) )
    *  The tag value still can be used for "standard" purposes. (But you will need a (very) litte
          change in your code, sorry ;( )
    *  Its faster.
    *  Handling is easier/more straight forward. (Only two simple digits in tag value to enter, not four)
    *  Prepared for handling 'Lines', too. (Not done by me - I don't use lines ;) )
    *  A var naming convention is used, so code is easier to read/modify.
    *  Demo and description extended.
    *  ...

    All  -(C) Copyleft on 11/10/2003 - Light Templer (LiTe)


    Usage:      Simply put on every control you want to automaticly be resized by ucAutoResize
                three additional digits to the end of the TAG value.

                1 -  '|  as delimiter to the rest. (Can be changed - look for Const TAG_DELIMITER )
                2 -  'L  or 'R  or '-'
                3 -  'T  or 'B  or '-'

                e.g.:       'Normal tag value1 |LB'
                            'Val1, Val2, Val3|R-'
                            '|-T'

                'L  :       LEFT border of the control will be moved when form is resized.
                            Width of control will not be changed.

                'R  :       RIGHT border of the control will be moved when form is resized.
                            Width of control will be changed.


                'T  :       TOP border of the control will be moved when form is resized.
                            Height of control will not be changed.

                'B  :       BOTTOM border of the control will be moved when form is resized.
                            Height of control will be changed.

                '-  :       Don't change anything on this direction.


    Thats all! Simple and easy to remember!
    Have fun with it and kind regards.

    LiTe

