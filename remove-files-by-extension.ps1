

get-childitem "X:\PATH\" -include *.txt -recurse | foreach ($_) {remove-item $_.fullname}
