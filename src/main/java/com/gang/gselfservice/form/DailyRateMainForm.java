package com.gang.gselfservice.form;

import com.gang.gselfservice.config.ThreadPool;
import com.gang.gselfservice.service.DailyRateService;
import com.gang.gselfservice.task.SourceBackupTask;
import com.gang.gselfservice.utils.DateUtils;
import com.intellij.uiDesigner.core.GridConstraints;
import com.intellij.uiDesigner.core.GridLayoutManager;
import com.intellij.uiDesigner.core.Spacer;
import org.apache.commons.lang3.StringUtils;
import org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.prefs.Preferences;

public class DailyRateMainForm {
    private JButton generateButton;
    private JPanel homePanel;
    private JLabel sourceLabel;
    private JLabel outputLabel;
    private JPanel sourcePanel;
    private JPanel outputPanel;
    private JTextField sourceTextField;
    private JTextField outputTextField;
    private JButton sourceButton;
    private JButton outputButton;

    private static Preferences pref = Preferences.userRoot().node(DailyRateMainForm.class.getName());

    public DailyRateMainForm() {
    }

    /*
     * 打开文件
     */
    private static void showFileOpenDialog(Component parent, JTextField textField) {
        // 创建一个默认的文件选取器
        JFileChooser fileChooser = new JFileChooser();

        // 设置默认显示的文件夹为当前文件夹
//        fileChooser.setCurrentDirectory(new File("~"));
        String lastSourcePath = pref.get("lastSourcePath", "");
        if (StringUtils.isNotBlank(lastSourcePath)) {
            fileChooser.setCurrentDirectory(new File(lastSourcePath));
        }

        // 设置文件选择的模式（只选文件、只选文件夹、文件和文件均可选）
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

        // 设置是否允许多选
        fileChooser.setMultiSelectionEnabled(false);

        // 显示列表
        fileChooser.getActionMap().get("viewTypeDetails").actionPerformed(null);

        // 添加可用的文件过滤器（FileNameExtensionFilter 的第一个参数是描述, 后面是需要过滤的文件扩展名 可变参数）
        FileNameExtensionFilter extensionFilter = new FileNameExtensionFilter("Excel 2007/2003 (*.xlsx,*.xls)", "xlsx", "xls");
        fileChooser.addChoosableFileFilter(extensionFilter);

        // 设置默认使用的文件过滤器
        fileChooser.setFileFilter(extensionFilter);
        // 不允许选择所有文件
        fileChooser.setAcceptAllFileFilterUsed(false);

        // 打开文件选择框（线程将被阻塞, 直到选择框被关闭）
        int result = fileChooser.showOpenDialog(parent);

        if (result == JFileChooser.APPROVE_OPTION) {
            // 如果点击了"确定", 则获取选择的文件路径
            File file = fileChooser.getSelectedFile();

            // 如果允许选择多个文件, 则通过下面方法获取选择的所有文件
            // File[] files = fileChooser.getSelectedFiles();

            textField.setText(file.getAbsolutePath());

            pref.put("lastSourcePath", textField.getText());
        }
    }

    /*
     * 选择文件保存路径
     */
    private static void showFileSaveDialog(Component parent, JTextField textField) {
        // 创建一个默认的文件选取器
        JFileChooser fileChooser = new JFileChooser();

        // 设置打开文件选择框后默认输入的文件名
        String lastOutputPath = pref.get("lastOutputPath", "");
        if (StringUtils.isNotBlank(lastOutputPath)) {
            fileChooser.setSelectedFile(new File(lastOutputPath));
        } else {
            fileChooser.setSelectedFile(new File(DateUtils.getCurrentDate() + ".docx"));
        }

        // 显示列表
        fileChooser.getActionMap().get("viewTypeDetails").actionPerformed(null);

        // 打开文件选择框（线程将被阻塞, 直到选择框被关闭）
        int result = fileChooser.showSaveDialog(parent);

        if (result == JFileChooser.APPROVE_OPTION) {
            // 如果点击了"保存", 则获取选择的保存路径
            File file = fileChooser.getSelectedFile();
            textField.setText(file.getAbsolutePath());
            pref.put("lastOutputPath", textField.getText());
        }
    }

    public static void main(String[] args) {
        // 样式初始化
        try {
            UIManager.put("RootPane.setupButtonVisible", false);
            BeautyEyeLNFHelper.frameBorderStyle = BeautyEyeLNFHelper.FrameBorderStyle.generalNoTranslucencyShadow;
            BeautyEyeLNFHelper.debug = false;
            BeautyEyeLNFHelper.translucencyAtFrameInactive = false;
            BeautyEyeLNFHelper.launchBeautyEyeLNF();
        } catch (Exception e) {
        }

        // form的初始化
        DailyRateMainForm dailyRateMainForm = new DailyRateMainForm();
        dailyRateMainForm.init();

        // 创建frame
        JFrame frame = new JFrame("满意度日报工具V1.1");
        frame.setContentPane(dailyRateMainForm.homePanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        frame.pack();
        frame.setResizable(false);

        // 居中
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        Point p = ge.getCenterPoint();
        frame.setLocation(p.x - frame.getWidth() / 2, p.y - frame.getHeight() / 2);

        frame.setVisible(true);
    }

    private void init() {
        // 生成按钮
        generateButton.addActionListener(e -> {
            // 先校验
            if (StringUtils.isBlank(sourceTextField.getText()) || StringUtils.isBlank(outputTextField.getText())) {
                JOptionPane.showMessageDialog(homePanel, "请先选择文件！");
                return;
            } else {
                File sourceFile = new File(sourceTextField.getText());
                if (!sourceFile.exists()) {
                    JOptionPane.showMessageDialog(homePanel, "源文件不存在！");
                    return;
                }
                // 保存备份
                ThreadPool.getExecutorService().submit(new SourceBackupTask(sourceTextField.getText()));
            }

            generateButton.setText("请等待……");
            generateButton.setEnabled(false);
            sourceButton.setEnabled(false);
            outputButton.setEnabled(false);

            // 异步执行逻辑
            SwingWorker<String, Object> worker = new SwingWorker<String, Object>() {

                @Override
                public String doInBackground() throws Exception {
                    // 所有报错信息都往外抛，最后弹窗
                    DailyRateService dailyRateService = new DailyRateService();
                    dailyRateService.dealDailyRateExcel(sourceTextField.getText());
                    dailyRateService.patchDailyRateSummary();
                    dailyRateService.exportToWord(outputTextField.getText());
                    return "生成";
                }

                @Override
                protected void done() {
                    try {
                        sourceButton.setEnabled(true);
                        outputButton.setEnabled(true);
                        generateButton.setEnabled(true);
                        generateButton.setText(get());
                        if (0 == JOptionPane.showConfirmDialog(homePanel, "是否直接打开文件?", "完成！",
                                JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE)) {
                            File file = new File(outputTextField.getText());
                            Desktop.getDesktop().open(file);
                        }
                    } catch (Exception exception) {
                        sourceButton.setEnabled(true);
                        outputButton.setEnabled(true);
                        generateButton.setEnabled(true);
                        generateButton.setText("生成");
                        JOptionPane.showMessageDialog(homePanel, exception.getMessage());
                    }
                }
            };
            worker.execute();
        });

        // 选择源文件
        sourceTextField.setEditable(false);
        sourceButton.addActionListener(e -> showFileOpenDialog(sourcePanel, sourceTextField));

        // 选择输出目录
        outputTextField.setEditable(false);
        outputButton.addActionListener(e -> showFileSaveDialog(outputPanel, outputTextField));

        // 上次选择
        String lastSourcePath = pref.get("lastSourcePath", "");
        String lastOutputPath = pref.get("lastOutputPath", "");

        if (StringUtils.isNotBlank(lastSourcePath)) {
            sourceTextField.setText(lastSourcePath);
        }
        if (StringUtils.isNotBlank(lastOutputPath)) {
            outputTextField.setText(lastOutputPath);
        }
    }

    {
// GUI initializer generated by IntelliJ IDEA GUI Designer
// >>> IMPORTANT!! <<<
// DO NOT EDIT OR ADD ANY CODE HERE!
        $$$setupUI$$$();
    }

    /**
     * Method generated by IntelliJ IDEA GUI Designer
     * >>> IMPORTANT!! <<<
     * DO NOT edit this method OR call it in your code!
     *
     * @noinspection ALL
     */
    private void $$$setupUI$$$() {
        homePanel = new JPanel();
        homePanel.setLayout(new GridLayoutManager(4, 2, new Insets(0, 0, 0, 0), -1, -1));
        generateButton = new JButton();
        generateButton.setText("生成");
        generateButton.setVerticalAlignment(0);
        homePanel.add(generateButton, new GridConstraints(3, 0, 1, 2, GridConstraints.ANCHOR_SOUTH, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        sourcePanel = new JPanel();
        sourcePanel.setLayout(new GridLayoutManager(2, 4, new Insets(0, 0, 0, 0), -1, -1));
        homePanel.add(sourcePanel, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        sourceLabel = new JLabel();
        sourceLabel.setText(" 源数据文件:");
        sourcePanel.add(sourceLabel, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        final Spacer spacer1 = new Spacer();
        sourcePanel.add(spacer1, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_VERTICAL, 1, GridConstraints.SIZEPOLICY_WANT_GROW, null, null, null, 0, false));
        sourceTextField = new JTextField();
        sourcePanel.add(sourceTextField, new GridConstraints(0, 1, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_FIXED, null, new Dimension(350, -1), null, 0, false));
        sourceButton = new JButton();
        sourceButton.setText("选择");
        sourcePanel.add(sourceButton, new GridConstraints(0, 2, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        final Spacer spacer2 = new Spacer();
        sourcePanel.add(spacer2, new GridConstraints(0, 3, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, 1, new Dimension(5, -1), null, null, 0, false));
        outputPanel = new JPanel();
        outputPanel.setLayout(new GridLayoutManager(2, 4, new Insets(0, 0, 0, 0), -1, -1));
        homePanel.add(outputPanel, new GridConstraints(2, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_BOTH, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, null, null, null, 0, false));
        outputLabel = new JLabel();
        outputLabel.setText(" 保存到文件:");
        outputPanel.add(outputLabel, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_NONE, GridConstraints.SIZEPOLICY_FIXED, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        final Spacer spacer3 = new Spacer();
        outputPanel.add(spacer3, new GridConstraints(1, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_VERTICAL, 1, GridConstraints.SIZEPOLICY_WANT_GROW, null, null, null, 0, false));
        outputTextField = new JTextField();
        outputPanel.add(outputTextField, new GridConstraints(0, 1, 1, 1, GridConstraints.ANCHOR_WEST, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, GridConstraints.SIZEPOLICY_FIXED, null, new Dimension(350, -1), null, 0, false));
        outputButton = new JButton();
        outputButton.setText("选择");
        outputPanel.add(outputButton, new GridConstraints(0, 2, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_CAN_SHRINK | GridConstraints.SIZEPOLICY_CAN_GROW, GridConstraints.SIZEPOLICY_FIXED, null, null, null, 0, false));
        final Spacer spacer4 = new Spacer();
        outputPanel.add(spacer4, new GridConstraints(0, 3, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_HORIZONTAL, GridConstraints.SIZEPOLICY_WANT_GROW, 1, new Dimension(5, -1), null, null, 0, false));
        final Spacer spacer5 = new Spacer();
        homePanel.add(spacer5, new GridConstraints(0, 0, 1, 1, GridConstraints.ANCHOR_CENTER, GridConstraints.FILL_VERTICAL, 1, GridConstraints.SIZEPOLICY_WANT_GROW, new Dimension(-1, 3), null, null, 0, false));
    }

    /**
     * @noinspection ALL
     */
    public JComponent $$$getRootComponent$$$() {
        return homePanel;
    }

}
