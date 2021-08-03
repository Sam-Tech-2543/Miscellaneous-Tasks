from rich.panel import Panel
from rich.console import Console
rc=Console()

#12 Hr to 24 Hr Time Converter
try:
    time=input("Enter the time(in 12 hour clock ex-> 1:39 pm):")
    a=time.split(" ")
    ampm=a[len(a)-1]
    t=a[0].split(":")
    hr=t[0]
    mi=t[1]
    m1=hr+":"+mi+" "+ampm

    if int(hr)<12 and int(mi)<=60 and hr>0 and ampm=="am":
        sorted=str(hr)+":"+str(mi)
        rc.print(Panel(f"[yellow]12 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]24 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif int(hr)<12 and int(mi)<=60 and hr>0  and ampm=="pm":
        sorted=str(int(hr)+12)+":"+str(mi)
        rc.print(Panel(f"[yellow]12 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]24 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif int(hr)==12 and int(mi)<=60 and hr>0  and ampm=="pm":
        sorted=str(hr)+":"+str(mi)
        rc.print(Panel(f"[yellow]12 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]24 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif int(hr)==12 and int(mi)<=60 and hr>0  and ampm=="am":
        sorted=str(int(hr)-12)+":"+str(mi)
        rc.print(Panel(f"[yellow]12 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]24 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    else:
        rc.print(Panel("[red]Invalid Inputs[/red]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
except:
    rc.print(Panel("[red]Invalid Inputs[/red]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))

#24 Hr to 12 Hr Converter
try:
    time=input("Enter the time(in 24 hour clock ex-> 21:3):")
    t=time.split(':')
    hr=int(t[0])
    mi=int(t[1])
    m1=str(hr)+":"+str(mi)
    if hr<12 and hr>0 and mi>=0 and mi<=60:
        sorted=str(hr)+":"+str(mi)+" am"
        rc.print(Panel(f"[yellow]24 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]12 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif hr>12 and hr<24 and mi>=0 and mi<=60:
        sorted=str(int(hr)-int(12))+":"+str(mi)+" pm"
        rc.print(Panel(f"[yellow]24 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]12 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif hr==0 and mi>=0 and mi<=60:
        sorted="12"+":"+str(mi)+" am"
        rc.print(Panel(f"[yellow]24 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]12 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif hr==12 and mi>=0 and mi<=60:
        sorted=str(hr)+":"+str(mi)+" pm"
        rc.print(Panel(f"[yellow]24 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]12 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    elif hr==24 and mi==0:
        sorted="12"+":"+str(mi)+" am"
        rc.print(Panel(f"[yellow]24 hour Time:-[/yellow][cyan] {m1}[/cyan]\n[yellow]12 hour Time:-[/yellow][cyan] {sorted}[/cyan]",title="[blue]12 hr to 24 hr Time Converter[/blue]"))
    else:
        rc.print(Panel("[red]Invalid Inputs[/red]",title="[blue]24 hr to 12 hr Time Converter[/blue]"))

except:
    rc.print(Panel("[red]Invalid Inputs[/red]",title="[blue]24 hr to 12 hr Time Converter[/blue]"))

#Sam Tech Op lel :-)