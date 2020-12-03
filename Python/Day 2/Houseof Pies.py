print("Welcome to the house of pies")
orderSelection ='y'
pies = ["Pecan", 
        "Apple Crisp", 
        "Bean", 
        "Banoffee",  
        "Black Bun", 
        "Blueberry", 
        "Buko", 
        "Burek",  
        "Tamale", 
        "Steak"]
pie_cart=[]
print("Here are our pies..")
for pieCount in pies:
    print(f' [{str(pies.index(pieCount))}] {pieCount}')


while orderSelection !='n':
    pieSelection = input("Which pie would you like?")
    pie_cart.append(pies[int(pieSelection)])

    print("Great well have that "+pies[int(pieSelection)-1] + " right out for you")

    orderSelection = input("Would you like to add another pie? (y/n)")

print(f'you order this many pies '+str(len(pie_cart )))


