def num_hash(num):
    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    if num < 26:
        return alpha[num-1]
    else:
        quotient, remainder = num // 26, num % 26
        if remainder == 0:
            if quotient == 1:
                return alpha[remainder-1]
            else:
                return num_hash(quotient-1) + alpha[remainder-1]
        else:
            return num_hash(quotient) + alpha[remainder-1]