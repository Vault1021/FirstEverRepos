Add-Type -AssemblyName system.drawing

# import the CSV and its data
$csv = Import-Csv "Images_to_rename.csv"

foreach($row in $csv){

#original file location - optional, the default is the directory the program runs from
# ex. c:\temp\
$original_filename_location = 'C:\Users\marco.reyes\OneDrive - Thomas Scientific\Documents\RenameImages\rename_images_differnt_names\original\'

#original file name that is read from csv file
$original_filename = $original_filename_location + $row.name.Substring($row.name.LastIndexOf("/") + 1)

#location where new file name should be copied to
$new_filename_location = 'C:\Users\marco.reyes\OneDrive - Thomas Scientific\Documents\RenameImages\rename_images_differnt_names\renamed\'

# test if it can find the filename
if(test-path $original_filename){

    if ([IO.Path]::GetExtension($original_filename) -eq '.png') 
    {
        #convert original image file to jpg
        $convertedfile = $original_filename
        $convertedfile = $convertedfile -replace '.png', '.jpg'

        # output info to user
        write-host "Converting '$original_filename' to '$convertedfile'"

        $image = [drawing.image]::FromFile($original_filename)
	
	# Create a new bitmap with the desired background color
	$backgroundColor = [System.Drawing.Color]::White
	$newBitmap = New-Object System.Drawing.Bitmap $image.Width, $image.Height
	$graphics = [System.Drawing.Graphics]::FromImage($newBitmap)
	$graphics.Clear($backgroundColor)

	# Draw the original image onto the new bitmap
	$graphics.DrawImage($image, 0, 0)

	# Save the new bitmap as JPEG
	$newBitmap.Save($convertedfile,[System.Drawing.Imaging.ImageFormat]::Jpeg)

	# Dispose the objects
	$graphics.Dispose()
	$newBitmap.Dispose()
	$image.Dispose()

	# Delete original png file now that it has been converted
            try {
                Remove-Item $original_filename
            } catch {
                Write-Warning "Failed to delete file '$original_filename': $_"
            }


        #set original file name to the converted file name for copy
        $original_filename = $convertedfile
    }



# get new filename from spreadsheet and prepend the desired location
$new_filename = $new_filename_location + $row.newname

# output info to user
write-host "Copying '$original_filename' to '$new_filename'"

# copy original filename to new filename
copy-item $original_filename -Destination $new_filename 
}
else{
# if not found, write error
write-warning "File '$original_filename' not found"
}
}

Read-Host -Prompt "Press Enter to exit:"
