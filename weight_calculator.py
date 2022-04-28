# Calculates the number of plates to add to the bar to reach a certain weight.
# Functions with any weight quantity, including micro-plates.
# weight_calc(int or float)
def weight_calc(total):
    BAR = 45
    # Note: Input weight * quantity loaded onto bar at a time in pairs
    # If 2 sets of a weight, add to list twice, ex: 10*2, 10*2 for 4 total 10s
    weights = [BAR, 45*2, 35*2, 25*2, 10*2, 10*2, 5*2, 2.5*2, 1*2, 0.75*2, 0.5*2, 0.25*2]
    weights_used = []
    # Function runs as long as the total is at least 0.5lbs
    while total >= 0.5:
        # If the total is greater than the first (heaviest weight) in the list of weights, subtract off the
        # weight of the first weight in the list and append it to weights_used, then remove it from the weights list
        if total >= weights[0]:
            total -= weights[0]
            weights_used.append(weights[0])

        del weights[0]

    del weights_used[0]  # removes the bar from the printout

    # Half the total for a set of 2 weight plates
    return [weights_used/2 for weights_used in weights_used]


