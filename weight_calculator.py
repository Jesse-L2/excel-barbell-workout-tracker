"""
weight_calculator.py
~~~~~~~~~~~~~~~~~~~~
This module calculates the number of plates to add to the bar to reach a certain weight.
Functions with any weight quantity, including micro-plates.
Input weight_calc(int or float)
Output is a list containing floating point numbers [float, float, ...]
"""


def weight_calc(total: float) -> list[float]:
    """
    Calculate weights needed to be added to barbell.
    
    :param total: Total weight on barbell
    :type total: int
    :raise TypeError: If n is not an int or float
    :return: A list of weights to use, all in one list
    :rtype: list
    
    Note: Input weight * quantity loaded onto bar at a time in pairs
    If 2 sets of a weight, add to list twice, ex: 10*2, 10*2 for 4 total 10s
    """

    BAR = 45
    weights = [BAR, 45 * 2, 35 * 2, 25 * 2, 10 * 2, 10 * 2, 5 * 2, 2.5 * 2, 1 * 2, 0.75 * 2, 0.5 * 2, 0.25 * 2]
    weights_used = []
    # Function runs as long as the total is at least 0.5lbs
    # TODO: change the while loop total amount to dynamically size to user's smallest weight
    while total >= 0.5:
        """If the total is greater than the first (the heaviest weight) in the list of weights, subtract off the
        weight of the first weight in the list and append it to weights_used, then remove it from the weights list"""
        if total >= weights[0]:
            total -= weights[0]
            weights_used.append(weights[0])
        # Always delete the first element of weights, regardless of whether it was added to weights_used
        del weights[0]

    del weights_used[0]  # removes the bar from the printout

    # Half the total for a set of 2 weight plates
    return [weights_used / 2 for weights_used in weights_used]


"""Notes for future improvements
1) Hold the weights list in the excel spreadsheet and pull those values into a dict with key = weight,
and value = quantity (in pairs)
2) Change the weight_calc function so that rather than deleting an element, it simply starts at the next index
Reasoning: it is not the best practice to delete items from a list as it is being iterated over
3) Would a deque be a better data structure to use? It would eliminate del weights[0] and weights[0]
"""
