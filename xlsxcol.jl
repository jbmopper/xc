using XLSX

function itocol(x::Int)
    if x <= 0
        throw(DomainError(x, "Numbers less than 1 are not valid Excel column indices."))
    end
    alphabet = 'A':1:'Z'
    l = length(alphabet)
    d, r = divrem(x, l)
    if r == 0
        r = l
        d -= 1
    end
    s = string(alphabet[r]) 
    while d > 0
        d, r = divrem(d, l)
        if r == 0
            r = l
            d -= 1
        end
        s = string(alphabet[r]) * s
    end
    s
end

#= 

julia> itocol(0)
ERROR: DomainError with 0:
Numbers less than 1 are not valid Excel column indices.
Stacktrace:
 [1] itocol(::Int64) at ./REPL[53]:3

julia> itocol(1)
"A"

julia> itocol(774)
"ACT"

julia> itocol(700)
"ZX"

=#
    


