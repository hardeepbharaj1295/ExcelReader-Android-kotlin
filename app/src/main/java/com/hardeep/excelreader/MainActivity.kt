package com.hardeep.excelreader

import android.Manifest
import android.content.Context
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Build
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.os.storage.StorageManager
import android.os.storage.StorageVolume
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import android.util.Log
import android.view.View
import android.widget.TextView
import java.io.InputStream
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import java.io.File
import java.io.FileInputStream

class MainActivity : AppCompatActivity() {

    lateinit var textView: TextView
    val TAG = "main"
    internal var PERMISSIONS =
        arrayOf(Manifest.permission.READ_EXTERNAL_STORAGE, android.Manifest.permission.WRITE_EXTERNAL_STORAGE)

    fun permissions():Boolean
    {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M)
        {
            if(checkSelfPermission(Manifest.permission.READ_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED)
            {
                requestPermissions(arrayOf(Manifest.permission.READ_EXTERNAL_STORAGE),1)
                return false
            }
            return true
        }
        return true
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        textView = findViewById(R.id.textview)

        val storageManager: StorageManager = this.getSystemService(Context.STORAGE_SERVICE) as StorageManager
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.N){
            val storageVolumeList:List<StorageVolume> =  storageManager.storageVolumes
            Log.e("Data",storageVolumeList.toString())
        } else {
            val storage:File=Environment.getExternalStorageDirectory()
            Log.e("Datas",storage.list().toString())
        }

        permissions()
//        readExcelFileFromAssets()
    }

    fun reader(v:View)
    {
        val intent = Intent(Intent.ACTION_GET_CONTENT)
        intent.setType("*/*")
        startActivityForResult(intent,1)
    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        when(requestCode){
            1->{
                val path = data!!.data.path
                readExcelFileFromAssets(path)
            }
        }
    }



    fun readExcelFileFromAssets(path:String)
    {
        try {
            Log.e("PAth",path)
            val data = path.replace("/file_share","")
            Log.e("Second",Environment.getExternalStorageDirectory().toString() + "/ssdpsgurugram/Book1.xls")
            val file = File(Environment.getExternalStorageDirectory().toString() + "/ssdpsgurugram/Book1.xls"
            )
            Log.e("Third",data)
            if (file.exists())
            {
                Log.e("Working","yes")
            }
            else
            {
                Log.e("Not","Done")
            }
            val input = FileInputStream(file)

            val myInput: InputStream
            // initialize asset manager
            val assetManager = assets
            //  open excel sheet
//            myInput = assetManager.open("Book1.xls")
            myInput = FileInputStream(data)
            // Create a POI File System object
            val myFileSystem = POIFSFileSystem(myInput)
            // Create a workbook using the File System
            val myWorkBook = HSSFWorkbook(myFileSystem)
            // Get the first sheet from workbook
            val mySheet = myWorkBook.getSheetAt(0)
            // We now need something to iterate through the cells.
            val rowIter = mySheet.rowIterator()
            var rowno = 0
            textView.append("\n")
            while (rowIter.hasNext()) {
                Log.e(TAG, " row no $rowno")
                val myRow = rowIter.next() as HSSFRow
                if (rowno != 0) {
                    val cellIter = myRow.cellIterator()
                    var colno = 0
                    var sno = ""
                    var date = ""
                    var det = ""
                    while (cellIter.hasNext()) {
                        val myCell = cellIter.next() as HSSFCell
                        if (colno == 0) {
                            sno = myCell.toString()
                        } else if (colno == 1) {
                            date = myCell.toString()
                        } else if (colno == 2) {
                            det = myCell.toString()
                        }
                        colno++
                        Log.e(TAG, " Index :" + myCell.columnIndex + " -- " + myCell.toString())
                    }
                    textView.append("$sno -- $date  -- $det\n")
                }
                rowno++
            }
        } catch (e: Exception) {
            Log.e(TAG, "error $e")
        }
    }
}
