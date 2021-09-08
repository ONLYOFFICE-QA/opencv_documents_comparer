import imageio
images = []

filenames = ['ictpol02_1_after_conv.png', 'ictpol02_1_before_conv.png']
for filename in filenames:
    images.append(imageio.imread(filename))
imageio.mimsave('movie.gif', images,  duration = 1)
