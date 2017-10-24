# Copyright Mentor Graphics Corporation 2003
#
#    All Rights Reserved.
#
# THIS WORK CONTAINS TRADE SECRET
# AND PROPRIETARY INFORMATION WHICH IS THE
# PROPERTY OF MENTOR GRAPHICS
# CORPORATION OR ITS LICENSORS AND IS
# SUBJECT TO LICENSE TERMS. 

fanout (direction in_out) (pin_type power) 
bus diagonal
unselect layer 3
route  10
clean  5
miter (style diagonal) 
write wire (include testpoint) C:\padspwr\files\previewpreroute.w
write routes (include testpoint) C:\padspwr\files\previewpreroute.rte