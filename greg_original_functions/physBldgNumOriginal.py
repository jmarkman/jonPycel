def physicalBuildingNum(final, caption):
	"""Identifies the number associated with the street1 and populates a colum with this number
	 if this column is Single Physical Building Number, it will copy Physical Building #
	
	 TODO test this more"""

	print 'PHYSICAL BULDING NUMBER IS RUNNING '
	print caption
	try:
		streetArr = final["Street 1"][0][:]
		streetArr.pop(0)
		numTracker = [caption]
		for val in streetArr:
			if len(val)>0:
				space = val.find(" ")
				dash = val.find("-")
				if dash != -1:
					num = val[:dash]
				else:
					num = val[:space]
				print "the numval %s" %(num)
				try:
					int(num)
					numTracker.append(str(num))
				except ValueError:
					pass
		if caption == 'Single Physical Building #':
			final[caption] = final['Physical Building #'][:]
		else:
			final[caption] = [numTracker]
		print final[caption]
	except:
		pass