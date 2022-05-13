
@echo  +-----------------------------------------------------------+
@echo  + Batch program to run the verification examples.           +
@echo  +-----------------------------------------------------------+
@echo.

@rem   ROHR2 Pfand hier eintragen
       PATH c:\SINETZ.38\SINETZW\

@echo To start
@echo.

@pause

@rem remove all existing result files

call clean_all.bat

@rem start calculations for all examples

       call sinetzw.exe  01_straight_pipe\01_straight_pipe.snp           /R

       call sinetzw.exe  02_straight_pipe_zeta\02_straight_pipe_zeta.snp /R

       call sinetzw.exe  03_straight_pipe_bend\03_straight_pipe_bend.snp /R

       call sinetzw.exe  04_pipe_reducer\04_straight_pipe_reducer.snp    /R

       call sinetzw.exe  05_pipe_tee\05_pipe+tee.snp                     /R

       call sinetzw.exe  06_straight_pipe+hight\06_pipe+hight_10m.snp    /R

       call sinetzw.exe  07_pipe_orifice\07_straight_pipe_orifice.snp    /R

       call sinetzw.exe  08_pump_given-curve\pumpe_kgs.snp               /R