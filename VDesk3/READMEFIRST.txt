The msghoo32.ocx freeware OCX from Mabry software is needed
by VirtualDesktop to recieve windows messages.

To begin, rename msghoo32.oc_ to msghoo32.ocx

Put this OCX into your windows\system directory then register it by going to
RUN and typing

regsvr32 c:\windows\system\msghoo32.ocx

Replace c:\windows\system\ with the path of your system dir.

Thanks,
Matt "The Code Monkey" Crowley
codemonkey04@cs.com
http://www.greenwave.org