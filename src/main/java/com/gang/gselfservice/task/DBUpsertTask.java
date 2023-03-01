package com.gang.gselfservice.task;

import com.gang.gselfservice.bo.DailyRateBO;
import com.gang.gselfservice.dao.RateDAO;
import org.apache.commons.collections4.CollectionUtils;

import java.sql.SQLException;
import java.util.List;
import java.util.concurrent.Callable;

public class DBUpsertTask implements Callable<Boolean> {

    private List<DailyRateBO> data;

    public DBUpsertTask(List<DailyRateBO> data) {
        this.data = data;
    }

    @Override
    public Boolean call() throws Exception {
        if (CollectionUtils.isNotEmpty(this.data)) {
            // 先建表
            RateDAO.createTable();

            // 如果有重复的日期，先删除
            String date = this.data.get(0).getDate();
            RateDAO.deleteByDate(date);

            // 循环插入数据，这里不做批量插入处理
            this.data.forEach(d -> {
                try {
                    RateDAO.addInfo(d);
                } catch (SQLException e) {
                    // ignore
                }
            });
        }

        return Boolean.TRUE;
    }
}
