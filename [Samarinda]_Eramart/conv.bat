@REM @Author: ichadhr
@REM @Date:   2018-07-11 10:50:16
@REM @Last Modified by:   richard.hari@live.com
@REM Modified time: 2018-10-18 14:04:06

pyuic5 -x gui.ui -o gui.py
pyrcc5 res_rc.qrc -o res_rc.py
pyuic5 -x gui2.ui -o gui2.py