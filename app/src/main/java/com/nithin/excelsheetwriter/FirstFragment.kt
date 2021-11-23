package com.nithin.excelsheetwriter

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import androidx.core.widget.addTextChangedListener
import androidx.fragment.app.Fragment
import androidx.navigation.fragment.findNavController
import com.nithin.excelsheetwriter.databinding.FragmentFirstBinding
import com.nithin.excelsheetwriter.utils.FileUtils
import com.nithin.excelsheetwriter.utils.SheetUtils

/**
 * A simple [Fragment] subclass as the default destination in the navigation.
 */
class FirstFragment : Fragment() {

    private var _binding: FragmentFirstBinding? = null

    // This property is only valid between onCreateView and
    // onDestroyView.
    private val binding get() = _binding!!

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View {
        _binding = FragmentFirstBinding.inflate(inflater, container, false)
        return binding.root
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)
        initViews()
        SheetUtils.setupSheet()
        binding.btnSubmit.setOnClickListener {
            findNavController().navigate(R.id.action_FirstFragment_to_SecondFragment)
        }
    }

    private fun initViews() {
        binding.apply {
            btnSubmitColumnName.setOnClickListener {
                SheetUtils.addColumnName(columnName = etColumnName.text.toString())
            }
            btnSubmitRowValue.addTextChangedListener {
                SheetUtils.addRowValue(rowValues = arrayListOf(etRowValue.text.toString()))
            }
            btnSubmit.setOnClickListener {
                FileUtils.createFile(context = it.context)
            }
        }
    }

    override fun onDestroyView() {
        super.onDestroyView()
        _binding = null
    }
}