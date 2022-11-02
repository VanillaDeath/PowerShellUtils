# Wilson's Progress Bar PowerShell class v1

# Usage:
#
# Include in your script:
# . .\Progress.class.ps1
#
# Construct an instance. Either:
# $foo = [Progress]::new() # Set properties using Setters later...
# $bar = [Progress]::new($num_items, $start_at, $Activity) # Activity is like a title of the operation.
# $foobar = [Progress]::new($num_items, $start_at, $Activity, $update_freq, $samples) # How often to update and how many updates to keep for time remaining estimation
#
# Start timer:
# $foo.Start()
#
# Call this method in your loop:
# $foo.SetAndShowProgress($current_item_num, $current_operation_label, $Force) # Your count, current item name, and whether to ignore $update_freq
# Setting $Force = $true on PowerShell versions older than Core (v6) will greatly slow down your loop for fast operations, since these versions will redraw the progress bar on every single update.
#  - Use only for operations for which an iteration generally takes longer than a split second, so an update for every iteration is useful.
# Setting $Force = $true on PowerShell Core has no effect, since Core has redraw-limiting built-in to Write-Progress
#
# Stop timer:
# $foo.Stop()

class Progress {
    # Properties:
    [int]$total = 1
    [int]$i = 0
    [decimal]$update = 0
    [decimal]$update_frequency = 1
    [Collections.Generic.List[array]]$Last10s
    [System.Diagnostics.Stopwatch]$StopWatch
    [string]$Activity = "Progress"
    [string]$Status = "Starting."
    [decimal]$PercentComplete
    [string]$CurrentOperation
    [int]$SecondsRemaining = -1
    [int]$num_diff
    [decimal]$time_diff
    [int]$num_remain
    [decimal]$ItemsPerSecond


    # Constructors:

    Progress( # Use all defaults Constructor, set values later...
    ) {
        $this.StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
        $this.Last10s = New-Object -TypeName Collections.Generic.List[array](10)
    }
    
    Progress( # Use default update frequency (1 sec) and num samples (10) Constructor.
        [int]$num_items,
        [int]$start_at,
        [string]$act
     ) {
        $this.total = $num_items
        $this.i = $start_at
        $this.update_frequency = 1
        $this.StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
        $this.Last10s = New-Object -TypeName Collections.Generic.List[array](10)
        $this.Activity = $act
     }    

    Progress( # Fully specified Constructor.
        [int]$num_items,
        [int]$start_at,
        [string]$act,
        [decimal]$update_freq,
        [int]$samples
    ) {
        $this.total = $num_items
        $this.i = $start_at
        $this.update_frequency = $update_freq
        $this.StopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
        $this.Activity = $act
        $this.Last10s = New-Object -TypeName Collections.Generic.List[array]($samples)
    }

    
    # Methods:

    [void]Start() {
        $this.StopWatch.Start()
    }

    [void]Stop() {
        $this.StopWatch.Stop()
    }

    [void]SetProgress([int]$count) {
        $this.i = $count
    }

    [void]SetCurrentOperation([string]$curr_oper) {
        $this.CurrentOperation = $curr_oper
    }

    [void]SetAndShowProgress($count, $curr_oper, $Force = $false) {
        if ($global:PSVersionTable.PSEdition -eq "Core" -or $Force -eq $true -or $this.update -eq 0 -or $this.StopWatch.Elapsed.TotalSeconds -ge $this.update + $this.update_frequency) {
            $this.i = $count
            $this.CurrentOperation = $curr_oper
            $this.ShowProgress()
        }
    }

    [void]ShowProgress() {
        if (-not $this.StopWatch.IsRunning) {
            $this.Start() # Start stopwatch if not running.
        }

        $elapsed = $this.StopWatch.Elapsed # Get elapsed info.
        $total_seconds = $elapsed.TotalSeconds # Get total elapsed seconds.

        if ($this.update -eq 0 -or $total_seconds -ge $this.update + $this.update_frequency) {
            while ($this.Last10s.Count -ge $this.Last10s.Capacity) { # Keep running this loop until last X updates list is no longer X items big (or more).
                $this.Last10s.RemoveAt(0) # Lop off X updates ago so first item in list is X-1 updates ago.
            }
            $this.Last10s.Add(@($this.i, $total_seconds)) # Add current item # and timestamp to last X updates progress list (back up to X updates).
            $refresh = $true
        } else {
            $refresh = $false
        }

        $this.num_diff = $this.i - $this.Last10s[0][0] # Number of items processed since X updates ago.
        $this.time_diff = $total_seconds - $this.Last10s[0][1] # Amount of time passed since X updates ago.
        $this.num_remain = $this.total - $this.i # Number of items remaining to process.
        $this.PercentComplete = ($this.i / $this.total) * 100 # Percent completed.

        if ($this.time_diff -gt 0) {
            $this.ItemsPerSecond = $this.num_diff / $this.time_diff # Approximate # items per second.
        } else {
            $this.ItemsPerSecond = 0 # Can't divide by 0, so make approx items/sec 0.
        }

        if ($refresh) {
            if ($this.num_diff -gt 0) {
                $this.SecondsRemaining = ($this.num_remain * $this.time_diff) / $this.num_diff # Remaining seconds based on how many items processed in last $time_diff seconds
            } else {
                $this.SecondsRemaining = -1 # Can't divide by 0, so make time remaining unknown (will hide it).
            }
        }

        $pad = $this.total.ToString().Length # How much padding to add to counts to keep text from shifting around.

        $this.Status = [Progress]::GetHMS($elapsed) + " elapsed.    " # HH:MM:SS elapsed.
        $this.Status += "[ " + ([math]::Round($this.PercentComplete)).ToString().PadLeft(3) + "% ]    " # Make percentage a whole number for status line.
        $this.Status += "[ " + $this.i.ToString().PadLeft($pad) + " / " + $this.total.ToString() + " ]    " # current / total
        $this.Status += "[ ~" + ([math]::Round($this.ItemsPerSecond)).ToString().PadLeft($pad) + " / sec ]" # ~ items / sec
        $num_remain_str = "[ " + $this.num_remain.ToString().PadLeft($pad) + " to go... ]" # Item countdown on right side.
        $this.Status += $num_remain_str.PadLeft([Progress]::GetWidth() - $this.Status.Length - 10) # Append item countdown to right.

        Write-Progress -Activity $this.Activity -Status $this.Status -PercentComplete $this.PercentComplete -SecondsRemaining $this.SecondsRemaining -CurrentOperation $this.CurrentOperation
        if ($refresh) {
            $this.update = $total_seconds # Set last update time to now.
        }
    }

    [void]SetActivity([string]$Act) {
        $this.Activity = $Act
    }

    [decimal]GetElapsedSeconds() {
        return $this.StopWatch.Elapsed.TotalSeconds
    }

    [decimal]GetUpdateFrequency() {
        return $this.update_frequency
    }

    [decimal]UpdatedAt() {
        return $this.update
    }

    [string]GetElapsed() {
        return [Progress]::GetHMS($this.StopWatch.Elapsed)
    }


    # Static Methods:

    static [int]GetWidth() { # Get # of character columns wide for this PowerShell window.
        return $global:host.UI.RawUI.WindowSize.width
    }

    static [void]SetWidth([int]$width) { # Resize width of PowerShell window to this many character columns.
        $global:host.UI.RawUI.WindowSize.width = $width
    }

    static [string]GetHMS([Object]$time_obj) { # Convert a StopWatch Elapsed object to an HH:MM:SS string.
        return "{0:00}" -f $time_obj.TotalHours + ":" + "{0:00}" -f $time_obj.Minutes + ":" + "{0:00}" -f $time_obj.Seconds # HH:MM:SS
    }
}