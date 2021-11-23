package com.nithin.excelsheetwriter

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Toast
import androidx.fragment.app.Fragment
import com.nithin.excelsheetwriter.databinding.FragmentSecondBinding
import com.nithin.excelsheetwriter.utils.FileUtils

/**
 * A simple [Fragment] subclass as the second destination in the navigation.
 */
class SecondFragment : Fragment() {

    private var _binding: FragmentSecondBinding? = null

    // This property is only valid between onCreateView and
    // onDestroyView.
    private val binding get() = _binding!!

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View {

        _binding = FragmentSecondBinding.inflate(inflater, container, false)
        return binding.root

    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)
        binding.buttonSecond.setOnClickListener {
            runWriteFile(it)
        }
    }

    private fun runWriteFile(view: View) {
        val isExcelGenerated: Boolean = FileUtils.exportDataIntoWorkbook(
            view.context, dataList = getDataList()
        )
        val strMessage: String = if (isExcelGenerated) {
            "Generated Successfully"
        } else {
            "Failed"
        }
        Toast.makeText(view.context, strMessage, Toast.LENGTH_SHORT).show()
    }


    private fun getDataList(): List<String> {
        return arrayListOf("Row 1, Row 2 , Row 3 , Row 4 , Row 5")
    }

    override fun onDestroyView() {
        super.onDestroyView()
        _binding = null
    }

}