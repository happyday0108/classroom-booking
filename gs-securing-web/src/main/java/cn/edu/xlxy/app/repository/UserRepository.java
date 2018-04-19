package cn.edu.xlxy.app.repository;

import org.springframework.data.repository.CrudRepository;

import cn.edu.xlxy.app.entity.User;

public interface  UserRepository  extends CrudRepository<User, Long> {

}
